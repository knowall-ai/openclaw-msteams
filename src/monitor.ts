/**
 * Inbound webhook server — receives Graph API notification POSTs.
 *
 * Handles:
 *   1. Subscription validation (responds with validationToken)
 *   2. Notification processing (validates clientState, dispatches to inbound)
 *   3. Subscription lifecycle (create + auto-renew timer)
 */

import type { Server } from "node:http";
import type { OpenClawConfig, RuntimeEnv } from "openclaw/plugin-sdk";
import { keepHttpServerTaskAlive } from "openclaw/plugin-sdk";
import { processNotification } from "./inbound.js";
import {
  buildSubscriptions,
  createSubscription,
  renewAll,
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

/**
 * Start the inbound webhook monitor.
 *
 * This is called by gateway.startAccount when OpenClaw starts.
 */
export async function startMonitor(opts: MonitorOpts): Promise<MonitorResult> {
  const channelCfg = opts.cfg.channels?.["msteams-user"] as MSTeamsUserConfig | undefined;

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

  const log = {
    info: (...args: unknown[]) => opts.runtime?.log?.(`[msteams-user] ${args.join(" ")}`),
    error: (...args: unknown[]) => opts.runtime?.error?.(`[msteams-user] ${args.join(" ")}`),
    debug: (...args: unknown[]) => opts.runtime?.log?.(`[msteams-user:debug] ${args.join(" ")}`),
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

          // Dispatch to OpenClaw via /hooks/wake
          await dispatchToOpenClaw(opts.cfg, message.text, message.sessionKey, log);
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
 * Dispatch an inbound message to OpenClaw via /hooks/wake.
 */
async function dispatchToOpenClaw(
  cfg: OpenClawConfig,
  text: string,
  sessionKey: string | null,
  log: { info: (...args: unknown[]) => void; error: (...args: unknown[]) => void },
): Promise<void> {
  // Try to find the hooks token from environment
  const hooksToken = process.env.HOOKS_TOKEN;
  if (!hooksToken) {
    log.error("HOOKS_TOKEN not set — cannot dispatch to OpenClaw");
    return;
  }

  const openclawUrl = process.env.OPENCLAW_URL ?? "http://localhost:3000";

  const payload: Record<string, unknown> = {
    text: text.slice(0, 2000),
    mode: "now",
  };
  if (sessionKey) {
    payload.sessionKey = sessionKey;
  }

  try {
    const resp = await fetch(`${openclawUrl}/hooks/wake`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${hooksToken}`,
      },
      body: JSON.stringify(payload),
    });

    if (resp.ok) {
      log.info(`Dispatched to OpenClaw${sessionKey ? ` [${sessionKey}]` : ""}`);
    } else {
      log.error(`Wake failed: ${resp.status} — ${await resp.text()}`);
    }
  } catch (err) {
    log.error("Failed to dispatch to OpenClaw:", err);
  }
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
