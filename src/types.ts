/**
 * Shared types for the msteams-user channel plugin.
 */

/** Resolved credentials for MS Graph API with delegated auth. */
export type MSTeamsUserCredentials = {
  clientId: string;
  tenantId: string;
  /** Optional — only needed if using confidential client flow. */
  clientSecret?: string;
  /** User's Azure AD object ID (for subscription resource paths). */
  userId?: string;
};

/** Channel config shape under channels.msteams-user in openclaw.json. */
export type MSTeamsUserConfig = {
  enabled?: boolean;
  clientId?: string;
  clientSecret?: string | { $secret: string };
  tenantId?: string;
  userId?: string;
  webhook?: {
    port?: number;
    clientState?: string;
    path?: string;
  };
  dmPolicy?: "open" | "allowlist" | "pairing";
  allowFrom?: Array<string | number>;
  groupAllowFrom?: Array<string | number>;
  defaultTo?: string;
};

/** Resolved account info for the config adapter. */
export type ResolvedMSTeamsUserAccount = {
  accountId: string;
  enabled: boolean;
  configured: boolean;
};

/** Result from sending a message via Graph API. */
export type SendMessageResult = {
  messageId: string;
  conversationId: string;
};

/** Graph API chat notification resource data. */
export type GraphNotification = {
  subscriptionId: string;
  changeType: string;
  resource: string;
  resourceData?: {
    id?: string;
    "@odata.type"?: string;
    "@odata.id"?: string;
  };
  clientState?: string;
  tenantId?: string;
};

/** Subscription definition for Graph webhook. */
export type SubscriptionDef = {
  name: string;
  resource: string;
  changeType: string;
  expirationMinutes: number;
  authType?: "delegated" | "application";
};
