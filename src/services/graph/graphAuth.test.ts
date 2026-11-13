import { describe, it, expect, beforeEach, vi } from "vitest";

vi.mock("../db/accounts", () => ({
  getAccount: vi.fn(),
  updateAccountAllTokens: vi.fn(),
  updateAccountTokens: vi.fn(),
}));
vi.mock("../oauth/oauthFlow", () => ({ refreshProviderToken: vi.fn() }));

import { getAccount, updateAccountAllTokens, updateAccountTokens } from "../db/accounts";
import { refreshProviderToken } from "../oauth/oauthFlow";
import { clearGraphTokenCache, getGraphAccessToken } from "./graphAuth";
import { createMockGmailAccount } from "@/test/mocks/entities.mock";

const nowSeconds = () => Math.floor(Date.now() / 1000);

function graphAccount(overrides = {}) {
  return createMockGmailAccount({
    id: "acc-graph",
    provider: "graph",
    access_token: "stored-access",
    refresh_token: "stored-refresh",
    token_expires_at: nowSeconds() + 3600,
    oauth_provider: "microsoft_graph",
    oauth_client_id: "client-1",
    oauth_tenant: "tenant-1",
    ...overrides,
  });
}

describe("getGraphAccessToken", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    clearGraphTokenCache();
  });

  it("returns the stored token while it is still valid", async () => {
    vi.mocked(getAccount).mockResolvedValue(graphAccount());

    await expect(getGraphAccessToken("acc-graph")).resolves.toBe("stored-access");
    await expect(getGraphAccessToken("acc-graph")).resolves.toBe("stored-access");

    expect(refreshProviderToken).not.toHaveBeenCalled();
    expect(getAccount).toHaveBeenCalledTimes(1); // second call served from cache
  });

  it("refreshes an expiring token against the account's tenant and stores the rotated refresh token", async () => {
    vi.mocked(getAccount).mockResolvedValue(graphAccount({ token_expires_at: nowSeconds() + 60 }));
    vi.mocked(refreshProviderToken).mockResolvedValue({
      access_token: "new-access",
      refresh_token: "new-refresh",
      expires_in: 3600,
      token_type: "Bearer",
    });

    await expect(getGraphAccessToken("acc-graph")).resolves.toBe("new-access");

    const [provider, refreshToken, clientId] = vi.mocked(refreshProviderToken).mock.calls[0]!;
    expect(provider.tokenUrl).toBe("https://login.microsoftonline.com/tenant-1/oauth2/v2.0/token");
    expect(refreshToken).toBe("stored-refresh");
    expect(clientId).toBe("client-1");
    expect(updateAccountAllTokens).toHaveBeenCalledWith("acc-graph", "new-access", "new-refresh", expect.any(Number));
  });

  it("keeps the old refresh token when none is returned", async () => {
    vi.mocked(getAccount).mockResolvedValue(graphAccount({ token_expires_at: 0 }));
    vi.mocked(refreshProviderToken).mockResolvedValue({ access_token: "new-access", expires_in: 3600, token_type: "Bearer" });

    await getGraphAccessToken("acc-graph");

    expect(updateAccountTokens).toHaveBeenCalledWith("acc-graph", "new-access", expect.any(Number));
    expect(updateAccountAllTokens).not.toHaveBeenCalled();
  });

  it("forces a refresh when asked", async () => {
    vi.mocked(getAccount).mockResolvedValue(graphAccount());
    vi.mocked(refreshProviderToken).mockResolvedValue({ access_token: "forced", expires_in: 3600, token_type: "Bearer" });

    await expect(getGraphAccessToken("acc-graph", true)).resolves.toBe("forced");
  });

  it("shares one refresh between concurrent callers", async () => {
    vi.mocked(getAccount).mockResolvedValue(graphAccount({ token_expires_at: 0 }));
    vi.mocked(refreshProviderToken).mockResolvedValue({ access_token: "shared", expires_in: 3600, token_type: "Bearer" });

    const [a, b] = await Promise.all([getGraphAccessToken("acc-graph"), getGraphAccessToken("acc-graph")]);

    expect(a).toBe("shared");
    expect(b).toBe("shared");
    expect(refreshProviderToken).toHaveBeenCalledTimes(1);
  });

  it("fails clearly when the account has no refresh token", async () => {
    vi.mocked(getAccount).mockResolvedValue(graphAccount({ refresh_token: null }));
    await expect(getGraphAccessToken("acc-graph")).rejects.toThrow(/sign in again/);
  });
});
