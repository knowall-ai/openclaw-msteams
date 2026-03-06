/**
 * Inbound notification processing — transforms Graph webhook notifications
 * into OpenClaw inbound messages.
 *
 * Reuses enrichment patterns from src/graph-adapter/index.js:
 *   - Fetches sender identity (displayName, email) from Graph API
 *   - Fetches recent message context
 *   - Self-message detection and filtering
 *   - Session routing: person:<email> for 1:1, group:<chatId> for groups
 */

import { getReadToken } from "./auth.js";
import type { MSTeamsUserCredentials, GraphNotification } from "./types.js";

export type InboundMessage = {
  /** Context text for the agent. */
  text: string;
  /** Session routing key (person:<email>, group:<chatId>). */
  sessionKey: string | null;
  /** Sender email (for allowlist matching). */
  senderEmail: string | null;
  /** Sender display name. */
  senderName: string | null;
  /** Chat ID. */
  chatId: string | null;
};

/** Debounce tracking — prevents self-notification loops. */
const recentWakes = new Map<string, number>();
const DEBOUNCE_MS = 10_000;

/**
 * Process a Graph notification into an inbound message for OpenClaw.
 *
 * Returns null if the notification should be ignored (self-message, debounced, etc.).
 */
export async function processNotification(
  notification: GraphNotification,
  creds: MSTeamsUserCredentials,
  ignoreUserId?: string,
): Promise<InboundMessage | null> {
  const { resource, changeType } = notification;

  // Only handle new messages
  if (changeType !== "created") {
    return null;
  }

  // Only handle Teams chat messages
  if (!resource.includes("chats") || !resource.includes("messages")) {
    return null;
  }

  // Extract chatId and messageId from the resource path
  // Format: chats('chatId')/messages('messageId') or chats/chatId/messages/messageId
  const chatIdMatch =
    resource.match(/chats\('([^']+)'\)/) || resource.match(/chats\/([^/]+)/);
  const msgIdMatch =
    resource.match(/messages\('([^']+)'\)/) || resource.match(/messages\/([^/]+)/);

  const chatId = chatIdMatch?.[1] ?? null;
  const messageId = msgIdMatch?.[1] ?? null;

  if (!chatId || !messageId) {
    return null;
  }

  // Filter out meeting chats (noisy group chats from Teams meetings)
  if (chatId.includes("meeting_")) {
    return null;
  }

  // Debounce: skip if we recently processed this chat
  const now = Date.now();
  const lastWake = recentWakes.get(chatId);
  if (lastWake && now - lastWake < DEBOUNCE_MS) {
    return null;
  }
  recentWakes.set(chatId, now);

  // Clean up old debounce entries
  for (const [id, ts] of recentWakes.entries()) {
    if (now - ts > 60_000) {
      recentWakes.delete(id);
    }
  }

  // Fetch sender info and chat type in parallel
  const token = await getReadToken(creds);
  if (!token) {
    // No token — return generic message
    return {
      text: "New Teams message received. Use the Teams skill to read and respond.",
      sessionKey: null,
      senderEmail: null,
      senderName: null,
      chatId,
    };
  }

  const [senderInfo, chatType, recentMessages] = await Promise.all([
    fetchSenderInfo(token, chatId, messageId),
    fetchChatType(token, chatId),
    fetchRecentMessages(token, chatId, 10),
  ]);

  // Filter self-messages
  if (ignoreUserId && senderInfo?.senderId === ignoreUserId) {
    return null;
  }

  // Derive session key
  let sessionKey: string | null = null;
  if (chatType === "oneOnOne" && senderInfo?.senderEmail) {
    sessionKey = `person:${senderInfo.senderEmail.toLowerCase()}`;
  } else if (chatType === "group" && chatId) {
    sessionKey = `group:${chatId}`;
  }

  // Build context message
  const chatDesc = chatType === "oneOnOne" ? "1-2-1 chat" : "group chat";
  const senderDesc =
    senderInfo?.senderName && senderInfo?.senderEmail
      ? `${senderInfo.senderName} (${senderInfo.senderEmail})`
      : "someone";

  let text: string;
  if (recentMessages.length > 0) {
    text =
      `New Teams message from ${senderDesc} in ${chatDesc}.\n\n` +
      `**Recent conversation (last ${recentMessages.length} messages):**\n` +
      recentMessages.join("\n") +
      "\n\n" +
      "Read the conversation above carefully. Respond to ALL unanswered messages. " +
      "If you don't have enough context, use the Teams skill to fetch more messages.";
  } else {
    text =
      `New Teams message from ${senderDesc} in ${chatDesc}. ` +
      "Use the Teams skill to fetch recent messages before responding.";
  }

  return {
    text,
    sessionKey,
    senderEmail: senderInfo?.senderEmail ?? null,
    senderName: senderInfo?.senderName ?? null,
    chatId,
  };
}

// --- Graph API helpers ---

type SenderInfo = {
  senderEmail: string | null;
  senderName: string | null;
  senderId: string | null;
};

async function fetchSenderInfo(
  token: string,
  chatId: string,
  messageId: string,
): Promise<SenderInfo | null> {
  try {
    const resp = await fetch(
      `https://graph.microsoft.com/v1.0/chats/${chatId}/messages/${messageId}?$select=from`,
      { headers: { Authorization: `Bearer ${token}` } },
    );
    if (!resp.ok) return null;

    const msg = (await resp.json()) as {
      from?: { user?: { id?: string; displayName?: string } };
    };
    const userId = msg.from?.user?.id ?? null;
    const senderName = msg.from?.user?.displayName ?? null;

    if (!userId) return { senderEmail: null, senderName, senderId: null };

    // Fetch user's email
    const userResp = await fetch(
      `https://graph.microsoft.com/v1.0/users/${userId}?$select=mail,userPrincipalName`,
      { headers: { Authorization: `Bearer ${token}` } },
    );
    if (!userResp.ok) {
      return { senderEmail: null, senderName, senderId: userId };
    }

    const user = (await userResp.json()) as {
      mail?: string;
      userPrincipalName?: string;
    };
    return {
      senderEmail: user.mail ?? user.userPrincipalName ?? null,
      senderName,
      senderId: userId,
    };
  } catch {
    return null;
  }
}

async function fetchChatType(
  token: string,
  chatId: string,
): Promise<string | null> {
  try {
    const resp = await fetch(
      `https://graph.microsoft.com/v1.0/chats/${chatId}?$select=chatType`,
      { headers: { Authorization: `Bearer ${token}` } },
    );
    if (!resp.ok) return null;

    const chat = (await resp.json()) as { chatType?: string };
    return chat.chatType ?? null;
  } catch {
    return null;
  }
}

async function fetchRecentMessages(
  token: string,
  chatId: string,
  count: number,
): Promise<string[]> {
  try {
    const resp = await fetch(
      `https://graph.microsoft.com/v1.0/chats/${chatId}/messages?$top=${count}&$orderby=createdDateTime desc&$select=from,body,createdDateTime`,
      { headers: { Authorization: `Bearer ${token}` } },
    );
    if (!resp.ok) return [];

    const data = (await resp.json()) as {
      value: Array<{
        from?: { user?: { displayName?: string } };
        body?: { content?: string };
        createdDateTime?: string;
      }>;
    };

    return (data.value ?? [])
      .filter((m) => m.from?.user)
      .reverse()
      .map((m) => {
        const name = m.from?.user?.displayName ?? "Unknown";
        let body = (m.body?.content ?? "").replace(/<[^>]+>/g, "").trim();
        if (body.length > 300) body = body.substring(0, 300) + "...";
        const time = m.createdDateTime
          ? new Date(m.createdDateTime).toLocaleTimeString("en-GB", {
              hour: "2-digit",
              minute: "2-digit",
            })
          : "";
        return `[${time}] ${name}: ${body}`;
      });
  } catch {
    return [];
  }
}
