/**
 * Graph API message sending — POST /chats/{chatId}/messages
 *
 * Uses delegated auth (as the user) rather than bot credentials.
 * This makes messages appear as sent by the user account (Sallie)
 * rather than a bot identity.
 */

import { getDelegatedToken } from "./auth.js";
import { resolveCredentials } from "./token.js";
import type { MSTeamsUserConfig, SendMessageResult } from "./types.js";

export type SendGraphMessageParams = {
  /** Channel config (for credentials). */
  channelCfg?: MSTeamsUserConfig;
  /** Chat ID or user target (user:email resolves to a chat). */
  to: string;
  /** Message text (HTML supported). */
  text: string;
  /** Optional media URL (uploaded separately). */
  mediaUrl?: string;
};

/**
 * Resolve a target to a chat ID.
 *
 * Targets can be:
 *   - A raw chat ID (19:xxx@thread.v2)
 *   - user:<email> — finds or creates a 1:1 chat with that user
 */
async function resolveChatId(
  to: string,
  token: string,
): Promise<string> {
  const trimmed = to.trim();

  // Already a chat ID
  if (trimmed.startsWith("19:") || trimmed.includes("@thread")) {
    return trimmed;
  }

  // user:<email> — find or create 1:1 chat
  const userMatch = trimmed.match(/^user:(.+)/i);
  if (userMatch) {
    const email = userMatch[1]!.trim();

    // First, look up the user's ID
    const userResp = await fetch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(email)}?$select=id`,
      { headers: { Authorization: `Bearer ${token}` } },
    );
    if (!userResp.ok) {
      throw new Error(`Failed to resolve user ${email}: ${userResp.status}`);
    }
    const user = (await userResp.json()) as { id: string };

    // Create or get existing 1:1 chat
    const chatResp = await fetch("https://graph.microsoft.com/v1.0/chats", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        chatType: "oneOnOne",
        members: [
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            roles: ["owner"],
            "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${user.id}')`,
          },
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            roles: ["owner"],
            // "me" — the authenticated user (Sallie)
            "user@odata.bind": "https://graph.microsoft.com/v1.0/me",
          },
        ],
      }),
    });
    if (!chatResp.ok) {
      const errBody = await chatResp.text();
      throw new Error(`Failed to create/get chat with ${email}: ${chatResp.status} — ${errBody}`);
    }
    const chat = (await chatResp.json()) as { id: string };
    return chat.id;
  }

  // Fall through — treat as raw chat ID
  return trimmed;
}

/**
 * Send a text message to a Teams chat via Graph API.
 */
export async function sendGraphMessage(
  params: SendGraphMessageParams,
): Promise<SendMessageResult> {
  const { channelCfg, to, text, mediaUrl } = params;

  const creds = resolveCredentials(channelCfg);
  if (!creds) {
    throw new Error("msteams-user credentials not configured");
  }

  const token = await getDelegatedToken(creds);
  if (!token) {
    throw new Error(
      "No delegated auth token. Run device-code login first " +
        "(openclaw channels login msteams-user)",
    );
  }

  const chatId = await resolveChatId(to, token);

  // Build message body
  let body: string;
  if (mediaUrl) {
    // Include media as a link in the message
    body = text ? `${text}\n\n📎 ${mediaUrl}` : `📎 ${mediaUrl}`;
  } else {
    body = text;
  }

  const resp = await fetch(
    `https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(chatId)}/messages`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        body: {
          contentType: "html",
          content: body,
        },
      }),
    },
  );

  if (!resp.ok) {
    const errBody = await resp.text();
    throw new Error(`Graph API send failed: ${resp.status} — ${errBody}`);
  }

  const result = (await resp.json()) as { id: string; chatId?: string };

  return {
    messageId: result.id,
    conversationId: chatId,
  };
}
