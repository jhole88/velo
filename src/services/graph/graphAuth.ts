import { getAccount, updateAccountAllTokens, updateAccountTokens } from "../db/accounts";
import { getMicrosoftGraphProvider } from "../oauth/providers";
import { refreshProviderToken } from "../oauth/oauthFlow";

/** Refresh this long before the access token actually expires. */
const REFRESH_BUFFER_MS = 5 * 60 * 1000;

interface CachedToken {
  accessToken: string;
  expiresAtMs: number;
}

const tokenCache = new Map<string, CachedToken>();
const inFlightRefresh = new Map<string, Promise<string>>();

/**
 * Get a valid Microsoft Graph access token for the account, refreshing it
 * when it is close to expiry (or when `forceRefresh` is set after a 401).
 *
 * Microsoft rotates refresh tokens, so a refresh persists both tokens.
 * Concurrent callers share a single in-flight refresh.
 */
export async function getGraphAccessToken(
  accountId: string,
  forceRefresh = false,
): Promise<string> {
  const cached = tokenCache.get(accountId);
  if (!forceRefresh && cached && cached.expiresAtMs - Date.now() > REFRESH_BUFFER_MS) {
    return cached.accessToken;
  }

  const pending = inFlightRefresh.get(accountId);
  if (pending) return pending;

  const promise = loadOrRefresh(accountId, forceRefresh).finally(() => {
    inFlightRefresh.delete(accountId);
  });
  inFlightRefresh.set(accountId, promise);
  return promise;
}

async function loadOrRefresh(accountId: string, forceRefresh: boolean): Promise<string> {
  const account = await getAccount(accountId);
  if (!account) throw new Error(`Account ${accountId} not found`);
  if (!account.refresh_token) {
    throw new Error(`Microsoft 365 account ${account.email} has no refresh token — sign in again.`);
  }
  if (!account.oauth_client_id) {
    throw new Error(`Microsoft 365 account ${account.email} has no application (client) ID.`);
  }

  const expiresAtMs = (account.token_expires_at ?? 0) * 1000; // DB stores seconds
  if (!forceRefresh && account.access_token && expiresAtMs - Date.now() > REFRESH_BUFFER_MS) {
    tokenCache.set(accountId, { accessToken: account.access_token, expiresAtMs });
    return account.access_token;
  }

  const provider = getMicrosoftGraphProvider(account.oauth_tenant);
  const tokens = await refreshProviderToken(
    provider,
    account.refresh_token,
    account.oauth_client_id,
    account.oauth_client_secret ?? undefined,
  );

  const newExpiresAt = Math.floor(Date.now() / 1000) + tokens.expires_in;
  if (tokens.refresh_token) {
    await updateAccountAllTokens(accountId, tokens.access_token, tokens.refresh_token, newExpiresAt);
  } else {
    await updateAccountTokens(accountId, tokens.access_token, newExpiresAt);
  }

  tokenCache.set(accountId, {
    accessToken: tokens.access_token,
    expiresAtMs: newExpiresAt * 1000,
  });
  return tokens.access_token;
}

/** Drop cached tokens (account removed or re-authenticated). */
export function clearGraphTokenCache(accountId?: string): void {
  if (accountId) {
    tokenCache.delete(accountId);
    inFlightRefresh.delete(accountId);
  } else {
    tokenCache.clear();
    inFlightRefresh.clear();
  }
}
