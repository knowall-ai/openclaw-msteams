/**
 * MSAL delegated auth — device-code flow with persistent token cache.
 *
 * Reuses the same pattern as workspace/scripts/lib/delegated-auth.mjs:
 *   1. Device-code login (one-time interactive setup)
 *   2. Token cache persisted to disk (refresh token lasts ~90 days)
 *   3. Silent renewal on subsequent calls
 *
 * The token cache is stored at ~/.openclaw/msteams-user-token-cache.json
 * separately from the existing graph-delegated-cache.json to avoid conflicts.
 */

import { PublicClientApplication } from "@azure/msal-node";
import { readFileSync, writeFileSync, mkdirSync } from "fs";
import { join } from "path";
import { homedir } from "os";
import type { MSTeamsUserCredentials } from "./types.js";

const CACHE_DIR = join(homedir(), ".openclaw");
const CACHE_PATH = join(CACHE_DIR, "msteams-user-token-cache.json");

/** Scopes needed for Teams messaging via delegated auth. */
const SCOPES = [
  "Chat.ReadWrite",
  "ChatMessage.Send",
  "ChannelMessage.Send",
  "User.Read",
];

/** Scopes for reading (used by inbound enrichment). */
const READ_SCOPES = [
  "Chat.Read",
  "User.ReadBasic.All",
];

function createMsalApp(creds: MSTeamsUserCredentials): PublicClientApplication {
  const app = new PublicClientApplication({
    auth: {
      clientId: creds.clientId,
      authority: `https://login.microsoftonline.com/${creds.tenantId}`,
    },
  });

  // Load persisted cache
  try {
    const data = readFileSync(CACHE_PATH, "utf8");
    app.getTokenCache().deserialize(data);
  } catch {
    // No cache yet — first login
  }

  return app;
}

function saveCache(app: PublicClientApplication): void {
  mkdirSync(CACHE_DIR, { recursive: true });
  writeFileSync(CACHE_PATH, app.getTokenCache().serialize());
}

/**
 * Interactive login via device-code flow.
 * Prints a URL and code for the user to visit in a browser.
 */
export async function login(
  creds: MSTeamsUserCredentials,
  onDeviceCode?: (message: string) => void,
): Promise<{ username: string }> {
  const app = createMsalApp(creds);

  const result = await app.acquireTokenByDeviceCode({
    scopes: [...SCOPES, ...READ_SCOPES],
    deviceCodeCallback: (response) => {
      if (onDeviceCode) {
        onDeviceCode(response.message);
      } else {
        console.log("\n" + response.message + "\n");
      }
    },
  });

  saveCache(app);

  return {
    username: result?.account?.username ?? "unknown",
  };
}

/**
 * Get a delegated access token using cached refresh token.
 * Returns null if no cache exists or refresh fails.
 */
export async function getDelegatedToken(
  creds: MSTeamsUserCredentials,
  scopes?: string[],
): Promise<string | null> {
  const app = createMsalApp(creds);

  const accounts = await app.getTokenCache().getAllAccounts();
  if (accounts.length === 0) {
    return null;
  }

  try {
    const result = await app.acquireTokenSilent({
      account: accounts[0]!,
      scopes: scopes ?? SCOPES,
    });
    // Persist updated cache (refresh token may have been rotated)
    saveCache(app);
    return result?.accessToken ?? null;
  } catch {
    return null;
  }
}

/**
 * Get a token suitable for reading (Chat.Read, User.ReadBasic.All).
 */
export async function getReadToken(creds: MSTeamsUserCredentials): Promise<string | null> {
  return getDelegatedToken(creds, [...READ_SCOPES, ...SCOPES]);
}

/**
 * Check if cached tokens exist (without attempting renewal).
 */
export async function hasCachedToken(creds: MSTeamsUserCredentials): Promise<boolean> {
  const app = createMsalApp(creds);
  const accounts = await app.getTokenCache().getAllAccounts();
  return accounts.length > 0;
}

/**
 * Clear all cached tokens.
 */
export async function logout(creds: MSTeamsUserCredentials): Promise<void> {
  const app = createMsalApp(creds);
  const accounts = await app.getTokenCache().getAllAccounts();
  for (const account of accounts) {
    await app.getTokenCache().removeAccount(account);
  }
  saveCache(app);
}
