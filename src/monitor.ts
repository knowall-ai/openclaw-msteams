/**
 * Inbound webhook server — receives Graph API notification POSTs and
 * dispatches through OpenClaw's channel reply pipeline.
 *
 * Previously dispatched via HTTP /hooks/wake. Now uses the SDK's
 * core.channel.reply.dispatchReplyFromConfig() for proper session
 * routing, reply delivery, and agent invocation.
 */

import type { Server } from "node:http";
import type { OpenClawConfig, ReplyPayload, RuntimeEnv } from "openclaw/plugin-sdk";
import {
  keepHttpServerTaskAlive,
  DEFAULT_ACCOUNT_ID,
  createReplyPrefixOptions,
} from "openclaw/plugin-sdk";
import { getRuntime } from "./runtime.js";
import { processNotification, type InboundMessage } from "./inbound.js";
import { sendGraphMessage } from "./send.js";
import {
  buildSubscriptions,
  createSubscription,
  startRenewalTimer,
  type SubscriptionManagerOpts,
} from "./subscriptions.js";
import { resolveCredentials } from "./token.js";
import type { GraphNotification, MSTeamsUserConfig } from "./types.js";

export type MonitorOpts = {
  cfg: OpenClawConfig;
  runtime?: RuntimeEnv;
  abortSignal?: AbortSignal;
};

export type MonitorResult = {
  app: unknown;
  shutdown: () => Promise<void>;
};

type Logger = {
  info: (...args: unknown[]) => void;
  error: (...args: unknown[]) => void;
  debug: (...args: unknown[]) => void;
};

function getChannelCfg(cfg: OpenClawConfig): MSTeamsUserConfig | undefined {
  return (cfg.channels as any)?.["msteams-user"];
}

/**
 * Dispatch an inbound message through the OpenClaw channel reply pipeline.
 *
 * Follows the same pattern as the official msteams/mattermost plugins:
 *   1. resolveAgentRoute → session key
 *   2. enqueueSystemEvent → heartbeat notification
 *   3. formatInboundEnvelope → formatted message body
 *   4. finalizeInboundContext → full context payload
 *   5. createReplyDispatcherWithTyping → reply delivery mechanism
 *   6. dispatchReplyFromConfig → invoke LLM and deliver reply
 */
async function dispatchInbound(
  message: InboundMessage,
  cfg: OpenClawConfig,
  channelCfg: MSTeamsUserConfig,
  log: Logger,
): Promise<void> {
  const core = getRuntime();
  const accountId = DEFAULT_ACCOUNT_ID;
  const chatId = message.chatId!;
  const chatType = message.chatType === "oneOnOne" ? ("direct" as const) : ("group" as const);

  // DM policy check
  const dmPolicy = channelCfg.dmPolicy ?? "pairing";
  if (chatType === "direct" && dmPolicy !== "open") {
    const allowFrom = (channelCfg.allowFrom ?? []).map((e) =>
      String(e).trim().toLowerCase(),
    );
    if (allowFrom.length > 0 && message.senderEmail) {
      if (!allowFrom.includes(message.senderEmail.toLowerCase())) {
        log.debug(`Sender ${message.senderEmail} not in allowlist, dropping`);
        return;
      }
    } else if (dmPolicy === "allowlist" && allowFrom.length === 0) {
      log.debug("DM policy is allowlist but no entries configured, dropping");
      return;
    }
  }

  // 1. Route resolution
  const peerId =
    chatType === "direct"
      ? (message.senderEmail ?? message.senderId ?? chatId)
      : chatId;
  const route = core.channel.routing.resolveAgentRoute({
    cfg,
    channel: "msteams-user",
    accountId,
    peer: { kind: chatType, id: peerId },
  });
  const sessionKey = route.sessionKey;

  // 2. Record activity
  core.channel.activity.record({
    channel: "msteams-user",
    accountId,
    direction: "inbound",
  });

  // 3. System event (visible in heartbeat)
  const preview = (message.rawText || "").replace(/\s+/g, " ").slice(0, 160);
  const senderLabel = message.senderName ?? message.senderEmail ?? "someone";
  const inboundLabel =
    chatType === "direct"
      ? `Teams DM from ${senderLabel}`
      : `Teams group message from ${senderLabel}`;
  core.system.enqueueSystemEvent(`${inboundLabel}: ${preview}`, {
    sessionKey,
    contextKey: `msteams-user:message:${chatId}:${message.messageId ?? "unknown"}`,
  });

  // 4. Format inbound envelope
  const fromLabel =
    chatType === "direct"
      ? message.senderName
        ? `${message.senderName} (${message.senderEmail ?? ""})`
        : (message.senderEmail ?? "unknown")
      : `${senderLabel} in group`;
  const body = core.channel.reply.formatInboundEnvelope({
    channel: "Microsoft Teams",
    from: fromLabel,
    body: message.text,
    chatType,
    sender: {
      name: message.senderName ?? undefined,
      id: message.senderId ?? undefined,
    },
  });

  // 5. Finalize inbound context
  const to = chatId;
  const ctxPayload = core.channel.reply.finalizeInboundContext({
    Body: body,
    BodyForAgent: message.rawText || message.text,
    RawBody: message.rawText || "",
    From:
      chatType === "direct"
        ? `msteams-user:${message.senderEmail ?? message.senderId ?? chatId}`
        : `msteams-user:group:${chatId}`,
    To: to,
    SessionKey: sessionKey,
    AccountId: route.accountId,
    ChatType: chatType,
    ConversationLabel: fromLabel,
    SenderName: message.senderName ?? undefined,
    SenderId: message.senderId ?? undefined,
    Provider: "msteams-user" as const,
    Surface: "msteams-user" as const,
    MessageSid: message.messageId ?? undefined,
    OriginatingChannel: "msteams-user" as const,
    OriginatingTo: to,
  });

  // 6. Update last route for DMs (so replies route back to this channel)
  if (chatType === "direct") {
    const sessionCfg = (cfg as any).session;
    const storePath = core.channel.session.resolveStorePath(
      sessionCfg?.store,
      { agentId: route.agentId },
    );
    await core.channel.session.updateLastRoute({
      storePath,
      sessionKey: route.mainSessionKey,
      deliveryContext: {
        channel: "msteams-user",
        to,
        accountId: route.accountId,
      },
    });
  }

  // 7. Create reply dispatcher with deliver callback
  const textLimit = core.channel.text.resolveTextChunkLimit(
    cfg,
    "msteams-user",
    accountId,
    { fallbackLimit: 4000 },
  );

  const { onModelSelected, ...prefixOptions } = createReplyPrefixOptions({
    cfg,
    agentId: route.agentId,
    channel: "msteams-user",
    accountId,
  });

  const { dispatcher, replyOptions, markDispatchIdle } =
    core.channel.reply.createReplyDispatcherWithTyping({
      ...prefixOptions,
      humanDelay: core.channel.reply.resolveHumanDelayConfig(cfg, route.agentId),
      deliver: async (payload: ReplyPayload) => {
        const text = payload.text ?? "";
        if (!text.trim()) return;

        const chunkMode = core.channel.text.resolveChunkMode(
          cfg,
          "msteams-user",
          accountId,
        );
        const chunks = core.channel.text.chunkMarkdownTextWithMode(
          text,
          textLimit,
          chunkMode,
        );
        for (const chunk of chunks.length > 0 ? chunks : [text]) {
          if (!chunk) continue;
          await sendGraphMessage({
            channelCfg,
            to,
            text: chunk,
          });
        }
        log.info(`Delivered reply to ${to}`);
      },
      onError: (err: unknown, info: { kind: string }) => {
        log.error(`Reply delivery failed (${info.kind}):`, err);
      },
    });

  // 8. Dispatch through LLM and deliver reply
  await core.channel.reply.withReplyDispatcher({
    dispatcher,
    onSettled: () => markDispatchIdle(),
    run: () =>
      core.channel.reply.dispatchReplyFromConfig({
        ctx: ctxPayload,
        cfg,
        dispatcher,
        replyOptions: {
          ...replyOptions,
          onModelSelected,
        },
      }),
  });
}

/**
 * Start the inbound webhook monitor.
 *
 * Called by gateway.startAccount when OpenClaw starts.
 */
export async function startMonitor(opts: MonitorOpts): Promise<MonitorResult> {
  const channelCfg = getChannelCfg(opts.cfg);

  if (!channelCfg?.enabled) {
    return { app: null, shutdown: async () => {} };
  }

  const creds = resolveCredentials(channelCfg);
  if (!creds) {
    opts.runtime?.error?.("msteams-user: credentials not configured");
    return { app: null, shutdown: async () => {} };
  }

  const port = channelCfg.webhook?.port ?? 3978;
  const clientState = channelCfg.webhook?.clientState ?? "msteams-user-graph-sub";
  const webhookPath = channelCfg.webhook?.path ?? "/webhooks/graph";
  const ignoreUserId = creds.userId;

  const log: Logger = {
    info: (...args) => opts.runtime?.log?.(`[msteams-user] ${args.join(" ")}`),
    error: (...args) => opts.runtime?.error?.(`[msteams-user] ${args.join(" ")}`),
    debug: (...args) => opts.runtime?.log?.(`[msteams-user:debug] ${args.join(" ")}`),
  };

  log.info(`Starting monitor on port ${port}`);

  const express = await import("express");
  const app = express.default();
  app.use(express.json({ limit: "1mb" }));

  // Webhook endpoint — handles both validation and notifications
  app.all(webhookPath, async (req: any, res: any) => {
    // Subscription validation: Graph sends GET with ?validationToken=
    if (req.query.validationToken) {
      log.debug("Subscription validation request");
      res.set("Content-Type", "text/plain");
      res.status(200).send(req.query.validationToken);
      return;
    }

    // Notification: Graph sends POST with body.value[]
    if (req.method === "POST" && req.body?.value) {
      log.debug(`Received ${req.body.value.length} notification(s)`);
      // Must respond within 3 seconds
      res.status(202).send();

      // Process notifications asynchronously
      for (const notification of req.body.value as GraphNotification[]) {
        try {
          // Validate clientState
          if (notification.clientState !== clientState) {
            log.debug("Invalid clientState, ignoring");
            continue;
          }

          const message = await processNotification(notification, creds, ignoreUserId);
          if (!message) continue;

          // Dispatch through SDK channel pipeline
          await dispatchInbound(message, opts.cfg, channelCfg, log);
        } catch (err) {
          log.error("Notification processing error:", err);
        }
      }
      return;
    }

    res.status(400).send("Invalid request");
  });

  // Health check
  app.get("/health", (_req: any, res: any) => {
    res.json({ status: "ok", channel: "msteams-user" });
  });

  // Start HTTP server
  const httpServer: Server = app.listen(port);
  await new Promise<void>((resolve, reject) => {
    const onListening = () => {
      httpServer.off("error", onError);
      log.info(`Monitor listening on port ${port}`);
      resolve();
    };
    const onError = (err: unknown) => {
      httpServer.off("listening", onListening);
      log.error("Server start failed:", err);
      reject(err);
    };
    httpServer.once("listening", onListening);
    httpServer.once("error", onError);
  });

  // Create subscriptions and start renewal timer
  if (creds.userId) {
    const subOpts: SubscriptionManagerOpts = {
      creds,
      webhookUrl: `${getWebhookBaseUrl(opts.cfg, channelCfg)}${webhookPath}`,
      clientState,
      log,
    };

    const subDefs = buildSubscriptions(creds.userId);

    // Create initial subscriptions
    for (const def of subDefs) {
      await createSubscription(subOpts, def);
    }

    // Start auto-renewal timer (every 50 min)
    startRenewalTimer(subOpts, subDefs, opts.abortSignal);
  } else {
    log.info("No userId configured — skipping subscription creation (use external subscriptions)");
  }

  const shutdown = async () => {
    log.info("Shutting down monitor");
    return new Promise<void>((resolve) => {
      httpServer.close(() => resolve());
    });
  };

  // Keep the task alive until abort
  await keepHttpServerTaskAlive({
    server: httpServer,
    abortSignal: opts.abortSignal,
    onAbort: shutdown,
  });

  return { app, shutdown };
}

/**
 * Build the public webhook base URL from config or environment.
 *
 * Priority:
 *   1. channels.msteams-user.webhook.url (explicit full URL)
 *   2. WEBHOOK_URL env var (used by existing graph-subscriptions)
 *   3. WEBHOOK_HOST env var + https://
 *   4. ui.hostname from config + https://
 */
function getWebhookBaseUrl(cfg: OpenClawConfig, channelCfg?: MSTeamsUserConfig): string {
  // Explicit webhook URL from channel config
  const configUrl = (channelCfg?.webhook as any)?.url;
  if (configUrl) return configUrl.replace(/\/$/, "");

  // WEBHOOK_URL from environment (same as graph-subscriptions uses)
  if (process.env.WEBHOOK_URL) {
    // WEBHOOK_URL may include a path — strip it to get just the base
    try {
      const parsed = new URL(process.env.WEBHOOK_URL);
      return `${parsed.protocol}//${parsed.host}`;
    } catch {
      return process.env.WEBHOOK_URL.replace(/\/$/, "");
    }
  }

  // WEBHOOK_HOST
  if (process.env.WEBHOOK_HOST) {
    return `https://${process.env.WEBHOOK_HOST}`;
  }

  // ui.hostname from config
  const hostname = (cfg as any).ui?.hostname;
  if (hostname) return `https://${hostname}`;

  return "https://localhost";
}
