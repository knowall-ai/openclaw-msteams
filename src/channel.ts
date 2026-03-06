/**
 * ChannelPlugin definition for msteams-user.
 *
 * Implements the same adapter interfaces as the official @openclaw/msteams
 * plugin, but uses Graph API with delegated auth instead of Bot Framework.
 */

import type { ChannelPlugin, OpenClawConfig } from "openclaw/plugin-sdk";
import {
  buildBaseChannelStatusSummary,
  createDefaultChannelRuntimeState,
  DEFAULT_ACCOUNT_ID,
  PAIRING_APPROVED_MESSAGE,
} from "openclaw/plugin-sdk";
import { outbound } from "./outbound.js";
import { resolveCredentials, hasConfiguredCredentials } from "./token.js";
import { getDelegatedToken, hasCachedToken } from "./auth.js";
import { sendGraphMessage } from "./send.js";
import type { ResolvedMSTeamsUserAccount, MSTeamsUserConfig } from "./types.js";

const meta = {
  id: "msteams-user",
  label: "Microsoft Teams (User)",
  selectionLabel: "Microsoft Teams (User Account)",
  docsPath: "/channels/msteams-user",
  docsLabel: "msteams-user",
  blurb: "Graph API; messages as your M365 user account.",
  aliases: ["teams-user"],
  order: 61,
} as const;

/** Helper to get channel config. */
function getChannelCfg(cfg: OpenClawConfig): MSTeamsUserConfig | undefined {
  return (cfg.channels as any)?.["msteams-user"];
}

export const msteamsUserPlugin: ChannelPlugin<ResolvedMSTeamsUserAccount> = {
  id: "msteams-user",

  meta: {
    ...meta,
    aliases: [...meta.aliases],
  },

  capabilities: {
    chatTypes: ["direct"],
    polls: false,
    threads: false,
    media: false,
  },

  agentPrompt: {
    messageToolHints: () => [
      "- MSTeams-User targeting: use `user:<email>` for DMs. Omit target to reply to current conversation.",
    ],
  },

  reload: { configPrefixes: ["channels.msteams-user"] },

  // --- Config adapter ---
  config: {
    listAccountIds: () => [DEFAULT_ACCOUNT_ID],

    resolveAccount: (cfg): ResolvedMSTeamsUserAccount => ({
      accountId: DEFAULT_ACCOUNT_ID,
      enabled: getChannelCfg(cfg)?.enabled !== false,
      configured: Boolean(resolveCredentials(getChannelCfg(cfg))),
    }),

    defaultAccountId: () => DEFAULT_ACCOUNT_ID,

    setAccountEnabled: ({ cfg, enabled }) => ({
      ...cfg,
      channels: {
        ...cfg.channels,
        "msteams-user": {
          ...getChannelCfg(cfg),
          enabled,
        },
      },
    }),

    deleteAccount: ({ cfg }) => {
      const next = { ...cfg } as OpenClawConfig;
      const nextChannels = { ...cfg.channels } as Record<string, unknown>;
      delete nextChannels["msteams-user"];
      if (Object.keys(nextChannels).length > 0) {
        (next as any).channels = nextChannels;
      } else {
        delete (next as any).channels;
      }
      return next;
    },

    isConfigured: (_account, cfg) => Boolean(resolveCredentials(getChannelCfg(cfg))),

    describeAccount: (account) => ({
      accountId: account.accountId,
      enabled: account.enabled,
      configured: account.configured,
    }),

    resolveAllowFrom: ({ cfg }) => getChannelCfg(cfg)?.allowFrom ?? [],

    formatAllowFrom: ({ allowFrom }) =>
      allowFrom
        .map((entry) => String(entry).trim())
        .filter(Boolean)
        .map((entry) => entry.toLowerCase()),

    resolveDefaultTo: ({ cfg }) => getChannelCfg(cfg)?.defaultTo?.trim() || undefined,
  },

  // --- Pairing adapter ---
  pairing: {
    idLabel: "email",
    normalizeAllowEntry: (entry) =>
      entry
        .replace(/^(msteams-user|msteams|user|teams-user):/i, "")
        .trim()
        .toLowerCase(),
    notifyApproval: async ({ cfg, id }) => {
      await sendGraphMessage({
        channelCfg: getChannelCfg(cfg),
        to: `user:${id}`,
        text: PAIRING_APPROVED_MESSAGE,
      });
    },
  },

  // --- Security adapter ---
  security: {
    collectWarnings: ({ cfg }) => {
      const channelCfg = getChannelCfg(cfg);
      if (!channelCfg) return [];

      const dmPolicy = channelCfg.dmPolicy ?? "pairing";
      if (dmPolicy === "open") {
        return [
          `- MS Teams (User): dmPolicy="open" allows any user to message. Set channels.msteams-user.dmPolicy="allowlist" + channels.msteams-user.allowFrom to restrict.`,
        ];
      }
      return [];
    },
  },

  // --- Setup adapter ---
  setup: {
    resolveAccountId: () => DEFAULT_ACCOUNT_ID,
    applyAccountConfig: ({ cfg }) => ({
      ...cfg,
      channels: {
        ...cfg.channels,
        "msteams-user": {
          ...getChannelCfg(cfg),
          enabled: true,
        },
      },
    }),
  },

  // --- Messaging adapter ---
  messaging: {
    normalizeTarget: (raw) => {
      const trimmed = raw.trim();
      if (!trimmed) return null;
      // Already prefixed
      if (/^user:/i.test(trimmed)) return trimmed.toLowerCase();
      // Email-like
      if (trimmed.includes("@") && !trimmed.includes(" ")) {
        return `user:${trimmed.toLowerCase()}`;
      }
      return trimmed;
    },
    targetResolver: {
      looksLikeId: (raw) => {
        const trimmed = raw.trim();
        if (!trimmed) return false;
        if (/^user:/i.test(trimmed)) return true;
        if (trimmed.includes("@") && !trimmed.includes(" ")) return true;
        return false;
      },
      hint: "<user:email>",
    },
  },

  // --- Directory adapter ---
  directory: {
    self: async () => null,

    listPeers: async ({ cfg, query, limit }) => {
      const q = query?.trim().toLowerCase() ?? "";
      const channelCfg = getChannelCfg(cfg);
      const ids = new Set<string>();

      for (const entry of channelCfg?.allowFrom ?? []) {
        const trimmed = String(entry).trim();
        if (trimmed && trimmed !== "*") {
          ids.add(trimmed.toLowerCase());
        }
      }

      return Array.from(ids)
        .map((raw) => {
          const cleaned = raw.replace(/^(msteams-user|user|teams-user):/i, "").trim();
          return cleaned.includes("@") ? `user:${cleaned}` : `user:${cleaned}`;
        })
        .filter((id) => (q ? id.toLowerCase().includes(q) : true))
        .slice(0, limit && limit > 0 ? limit : undefined)
        .map((id) => ({ kind: "user" as const, id }));
    },

    listGroups: async () => [],
  },

  // --- Outbound adapter ---
  outbound,

  // --- Status adapter ---
  status: {
    defaultRuntime: createDefaultChannelRuntimeState(DEFAULT_ACCOUNT_ID, { port: null }),

    buildChannelSummary: ({ snapshot }) => ({
      ...buildBaseChannelStatusSummary(snapshot),
      port: snapshot.port ?? null,
      probe: snapshot.probe,
    }),

    probeAccount: async ({ cfg }) => {
      const channelCfg = getChannelCfg(cfg);
      const creds = resolveCredentials(channelCfg);
      if (!creds) {
        return { ok: false, error: "missing credentials (clientId, tenantId)" };
      }

      try {
        const token = await getDelegatedToken(creds);
        if (!token) {
          const hasCache = await hasCachedToken(creds);
          return {
            ok: false,
            error: hasCache
              ? "token refresh failed — re-run device-code login"
              : "no cached token — run device-code login first",
          };
        }

        // Verify token by calling /me
        const resp = await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,mail", {
          headers: { Authorization: `Bearer ${token}` },
        });

        if (!resp.ok) {
          return { ok: false, error: `Graph /me returned ${resp.status}` };
        }

        const user = (await resp.json()) as { displayName?: string; mail?: string };
        return {
          ok: true,
          user: user.displayName ?? user.mail ?? "unknown",
        };
      } catch (err) {
        return { ok: false, error: String(err) };
      }
    },

    buildAccountSnapshot: ({ account, runtime, probe }) => ({
      accountId: account.accountId,
      enabled: account.enabled,
      configured: account.configured,
      running: runtime?.running ?? false,
      lastStartAt: runtime?.lastStartAt ?? null,
      lastStopAt: runtime?.lastStopAt ?? null,
      lastError: runtime?.lastError ?? null,
      port: runtime?.port ?? null,
      probe,
    }),
  },

  // --- Gateway adapter ---
  gateway: {
    startAccount: async (ctx) => {
      const { startMonitor } = await import("./monitor.js");
      const channelCfg = getChannelCfg(ctx.cfg);
      const port = channelCfg?.webhook?.port ?? 3978;
      ctx.setStatus({ accountId: ctx.accountId, port });
      ctx.log?.info?.(`msteams-user: starting monitor (port ${port})`);
      return startMonitor({
        cfg: ctx.cfg,
        runtime: ctx.runtime,
        abortSignal: ctx.abortSignal,
      });
    },
  },

  // --- Auth adapter ---
  auth: {
    login: async ({ cfg, runtime }) => {
      const { login } = await import("./auth.js");
      const channelCfg = getChannelCfg(cfg);
      const creds = resolveCredentials(channelCfg);
      if (!creds) {
        runtime.error?.("msteams-user: credentials not configured");
        return;
      }
      const result = await login(creds, (msg) => runtime.log?.(msg));
      runtime.log?.(`msteams-user: logged in as ${result.username}`);
    },
  },
};
