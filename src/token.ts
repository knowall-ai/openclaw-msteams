import type { MSTeamsUserConfig, MSTeamsUserCredentials } from "./types.js";

/**
 * Resolve a secret input value — handles both plain strings
 * and { $secret: "ENV_VAR" } references.
 */
function resolveSecretInput(value: unknown): string | undefined {
  if (typeof value === "string") {
    return value.trim() || undefined;
  }
  if (value && typeof value === "object" && "$secret" in value) {
    const envVar = (value as { $secret: string }).$secret;
    return process.env[envVar]?.trim() || undefined;
  }
  return undefined;
}

/**
 * Check whether the channel config has enough credentials configured.
 */
export function hasConfiguredCredentials(cfg?: MSTeamsUserConfig): boolean {
  return Boolean(resolveCredentials(cfg));
}

/**
 * Resolve credentials from channel config + environment variables.
 *
 * Priority: config values > environment variables.
 */
export function resolveCredentials(cfg?: MSTeamsUserConfig): MSTeamsUserCredentials | undefined {
  const clientId =
    cfg?.clientId?.trim() ||
    process.env.MSTEAMS_USER_CLIENT_ID?.trim() ||
    process.env.MS365_MCP_CLIENT_ID?.trim();

  const tenantId =
    cfg?.tenantId?.trim() ||
    process.env.MSTEAMS_USER_TENANT_ID?.trim() ||
    process.env.MS365_MCP_TENANT_ID?.trim();

  if (!clientId || !tenantId) {
    return undefined;
  }

  const clientSecret =
    resolveSecretInput(cfg?.clientSecret) ||
    process.env.MSTEAMS_USER_CLIENT_SECRET?.trim() ||
    process.env.MS365_MCP_CLIENT_SECRET?.trim();

  const userId =
    cfg?.userId?.trim() ||
    process.env.MSTEAMS_USER_USER_ID?.trim() ||
    process.env.MS365_USER_ID?.trim();

  return {
    clientId,
    tenantId,
    ...(clientSecret ? { clientSecret } : {}),
    ...(userId ? { userId } : {}),
  };
}
