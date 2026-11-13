export interface OAuthProviderConfig {
  id: string;
  name: string;
  authUrl: string;
  tokenUrl: string;
  scopes: string[];
  userInfoUrl?: string;
  /** Whether PKCE is required (Microsoft requires it, Yahoo supports it) */
  usePkce: boolean;
}

const providers: Record<string, OAuthProviderConfig> = {
  microsoft: {
    id: "microsoft",
    name: "Microsoft",
    authUrl:
      "https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize",
    tokenUrl:
      "https://login.microsoftonline.com/consumers/oauth2/v2.0/token",
    scopes: [
      "https://outlook.office.com/IMAP.AccessAsUser.All",
      "https://outlook.office.com/SMTP.Send",
      "offline_access",
      "openid",
      "profile",
      "email",
    ],
    userInfoUrl: undefined,
    usePkce: true,
  },
  yahoo: {
    id: "yahoo",
    name: "Yahoo",
    authUrl: "https://api.login.yahoo.com/oauth2/request_auth",
    tokenUrl: "https://api.login.yahoo.com/oauth2/get_token",
    scopes: ["mail-r", "mail-w", "openid", "sdps-r"],
    userInfoUrl: "https://api.login.yahoo.com/openid/v1/userinfo",
    usePkce: true,
  },
};

/** OAuth provider id used for Microsoft 365 work/school accounts via Microsoft Graph. */
export const MICROSOFT_GRAPH_PROVIDER_ID = "microsoft_graph";

/** Tenant segment used when no specific tenant is configured (any work/school account). */
export const DEFAULT_GRAPH_TENANT = "organizations";

/** Delegated Microsoft Graph permissions needed for mail and calendar. */
export const MICROSOFT_GRAPH_SCOPES = [
  "https://graph.microsoft.com/Mail.ReadWrite",
  "https://graph.microsoft.com/Mail.Send",
  "https://graph.microsoft.com/Calendars.ReadWrite",
  "https://graph.microsoft.com/MailboxSettings.Read",
  "https://graph.microsoft.com/User.Read",
  "offline_access",
  "openid",
  "profile",
  "email",
];

/** Tenant IDs are GUIDs or verified domains — reject anything else so it can't alter the URL. */
function normalizeTenant(tenant?: string | null): string {
  const trimmed = tenant?.trim();
  if (!trimmed) return DEFAULT_GRAPH_TENANT;
  if (!/^[A-Za-z0-9.-]+$/.test(trimmed)) {
    throw new Error("Invalid Microsoft tenant ID. Use the Directory (tenant) ID or your domain.");
  }
  return trimmed;
}

/**
 * Build the OAuth config for Microsoft Graph (Microsoft 365 / Exchange Online).
 * Single-tenant app registrations must sign in against their own tenant, so the
 * authority is built per account rather than hard-coded.
 */
export function getMicrosoftGraphProvider(tenant?: string | null): OAuthProviderConfig {
  const authority = `https://login.microsoftonline.com/${normalizeTenant(tenant)}/oauth2/v2.0`;
  return {
    id: MICROSOFT_GRAPH_PROVIDER_ID,
    name: "Microsoft 365",
    authUrl: `${authority}/authorize`,
    tokenUrl: `${authority}/token`,
    scopes: MICROSOFT_GRAPH_SCOPES,
    userInfoUrl: undefined,
    usePkce: true,
  };
}

/** True for both Microsoft flavours (personal IMAP OAuth and Microsoft Graph). */
export function isMicrosoftProvider(id: string): boolean {
  return id === "microsoft" || id === MICROSOFT_GRAPH_PROVIDER_ID;
}

export function getOAuthProvider(id: string): OAuthProviderConfig | null {
  if (id === MICROSOFT_GRAPH_PROVIDER_ID) return getMicrosoftGraphProvider();
  return providers[id] ?? null;
}

export function getAllOAuthProviders(): OAuthProviderConfig[] {
  return Object.values(providers);
}
