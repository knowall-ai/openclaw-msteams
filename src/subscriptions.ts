/**
 * Graph webhook subscription lifecycle — create, renew, delete.
 *
 * Reuses patterns from src/graph-subscriptions/manage.js:
 *   - Teams chat subs expire in 60 minutes (auto-renew timer)
 *   - Missing subscriptions are detected and recreated on renewal
 *   - Uses delegated auth for /chats/getAllMessages subscription
 */

import { getDelegatedToken } from "./auth.js";
import type { MSTeamsUserCredentials } from "./types.js";

export type SubscriptionConfig = {
  name: string;
  resource: string;
  changeType: string;
  expirationMinutes: number;
};

export type ActiveSubscription = {
  id: string;
  resource: string;
  changeType: string;
  expirationDateTime: string;
  notificationUrl: string;
};

export type SubscriptionManagerOpts = {
  creds: MSTeamsUserCredentials;
  webhookUrl: string;
  clientState: string;
  log?: {
    info: (...args: unknown[]) => void;
    error: (...args: unknown[]) => void;
    debug?: (...args: unknown[]) => void;
  };
};

/**
 * Build the subscription definitions for this user.
 */
export function buildSubscriptions(userId: string): SubscriptionConfig[] {
  return [
    {
      name: "teams-chat-messages",
      resource: `/users/${userId}/chats/getAllMessages`,
      changeType: "created",
      // Teams chat subs max out at 60 minutes
      expirationMinutes: 60,
    },
  ];
}

function getExpirationDateTime(minutes: number): string {
  const exp = new Date();
  exp.setMinutes(exp.getMinutes() + minutes);
  return exp.toISOString();
}

/**
 * Create a single Graph subscription.
 */
export async function createSubscription(
  opts: SubscriptionManagerOpts,
  subConfig: SubscriptionConfig,
): Promise<ActiveSubscription | null> {
  const token = await getDelegatedToken(opts.creds);
  if (!token) {
    opts.log?.error(`No delegated token for subscription: ${subConfig.name}`);
    return null;
  }

  const body = {
    changeType: subConfig.changeType,
    notificationUrl: opts.webhookUrl,
    resource: subConfig.resource,
    expirationDateTime: getExpirationDateTime(subConfig.expirationMinutes),
    clientState: opts.clientState,
  };

  opts.log?.info(`Creating subscription: ${subConfig.name} (${subConfig.resource})`);

  const resp = await fetch("https://graph.microsoft.com/v1.0/subscriptions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!resp.ok) {
    const errBody = await resp.text();
    opts.log?.error(`Failed to create subscription ${subConfig.name}: ${resp.status} — ${errBody}`);
    return null;
  }

  const result = (await resp.json()) as ActiveSubscription;
  opts.log?.info(`Subscription created: ${result.id} (expires ${result.expirationDateTime})`);
  return result;
}

/**
 * List all active subscriptions visible to the delegated user.
 */
export async function listSubscriptions(
  opts: SubscriptionManagerOpts,
): Promise<ActiveSubscription[]> {
  const token = await getDelegatedToken(opts.creds);
  if (!token) {
    return [];
  }

  const resp = await fetch("https://graph.microsoft.com/v1.0/subscriptions", {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!resp.ok) {
    opts.log?.error(`Failed to list subscriptions: ${resp.status}`);
    return [];
  }

  const data = (await resp.json()) as { value: ActiveSubscription[] };
  return data.value ?? [];
}

/**
 * Renew a subscription by ID.
 */
export async function renewSubscription(
  opts: SubscriptionManagerOpts,
  subscriptionId: string,
  expirationMinutes: number,
): Promise<boolean> {
  const token = await getDelegatedToken(opts.creds);
  if (!token) {
    return false;
  }

  const resp = await fetch(
    `https://graph.microsoft.com/v1.0/subscriptions/${subscriptionId}`,
    {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        expirationDateTime: getExpirationDateTime(expirationMinutes),
      }),
    },
  );

  if (!resp.ok) {
    opts.log?.error(`Failed to renew subscription ${subscriptionId}: ${resp.status}`);
    return false;
  }

  opts.log?.info(`Renewed subscription ${subscriptionId}`);
  return true;
}

/**
 * Delete a subscription by ID.
 */
export async function deleteSubscription(
  opts: SubscriptionManagerOpts,
  subscriptionId: string,
): Promise<boolean> {
  const token = await getDelegatedToken(opts.creds);
  if (!token) {
    return false;
  }

  const resp = await fetch(
    `https://graph.microsoft.com/v1.0/subscriptions/${subscriptionId}`,
    { method: "DELETE", headers: { Authorization: `Bearer ${token}` } },
  );

  if (!resp.ok) {
    opts.log?.error(`Failed to delete subscription ${subscriptionId}: ${resp.status}`);
    return false;
  }

  return true;
}

/**
 * Renew all existing subscriptions and recreate any missing ones.
 *
 * This is critical because Teams chat subs expire after 60 minutes,
 * so after a VM restart they will be gone.
 */
export async function renewAll(
  opts: SubscriptionManagerOpts,
  subscriptionDefs: SubscriptionConfig[],
): Promise<void> {
  const active = await listSubscriptions(opts);

  // Renew existing
  for (const sub of active) {
    const def = subscriptionDefs.find((d) =>
      sub.resource.includes(d.resource.split("/").pop()!),
    );
    const minutes = def?.expirationMinutes ?? 60;
    await renewSubscription(opts, sub.id, minutes);
  }

  // Recreate missing
  for (const def of subscriptionDefs) {
    const exists = active.some((sub) =>
      sub.resource.includes(def.resource.split("/").pop()!),
    );
    if (!exists) {
      opts.log?.info(`Missing subscription ${def.name} — recreating`);
      await createSubscription(opts, def);
    }
  }
}

/**
 * Start the subscription renewal timer.
 * Renews every 50 minutes (before the 60-minute expiry).
 */
export function startRenewalTimer(
  opts: SubscriptionManagerOpts,
  subscriptionDefs: SubscriptionConfig[],
  signal?: AbortSignal,
): void {
  const RENEWAL_INTERVAL_MS = 50 * 60 * 1000; // 50 minutes

  const timer = setInterval(async () => {
    try {
      await renewAll(opts, subscriptionDefs);
    } catch (err) {
      opts.log?.error("Subscription renewal failed:", err);
    }
  }, RENEWAL_INTERVAL_MS);

  signal?.addEventListener("abort", () => {
    clearInterval(timer);
  });
}
