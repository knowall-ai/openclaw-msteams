/**
 * ChannelOutboundAdapter — handles sending messages from OpenClaw to Teams.
 *
 * Uses Graph API POST /chats/{chatId}/messages with delegated auth,
 * so messages appear as sent by the user account (not a bot).
 */

import type { ChannelOutboundAdapter } from "openclaw/plugin-sdk";
import { getRuntime } from "./runtime.js";
import { sendGraphMessage } from "./send.js";

export const outbound: ChannelOutboundAdapter = {
  deliveryMode: "direct",

  chunker: (text, limit) => getRuntime().channel.text.chunkMarkdownText(text, limit),
  chunkerMode: "markdown",
  textChunkLimit: 4000,

  sendText: async ({ cfg, to, text }) => {
    const result = await sendGraphMessage({
      channelCfg: cfg.channels?.["msteams-user"],
      to,
      text,
    });
    return { channel: "msteams-user", ...result };
  },

  sendMedia: async ({ cfg, to, text, mediaUrl }) => {
    const result = await sendGraphMessage({
      channelCfg: cfg.channels?.["msteams-user"],
      to,
      text,
      mediaUrl,
    });
    return { channel: "msteams-user", ...result };
  },
};
