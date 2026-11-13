import { describe, it, expect } from "vitest";
import { getOAuthProvider, getAllOAuthProviders, getMicrosoftGraphProvider, isMicrosoftProvider } from "./providers";

describe("getOAuthProvider", () => {
  it("returns microsoft provider config", () => {
    const provider = getOAuthProvider("microsoft");
    expect(provider).not.toBeNull();
    expect(provider!.id).toBe("microsoft");
    expect(provider!.name).toBe("Microsoft");
    expect(provider!.authUrl).toContain("login.microsoftonline.com");
    expect(provider!.tokenUrl).toContain("login.microsoftonline.com");
    expect(provider!.scopes).toContain("https://outlook.office.com/IMAP.AccessAsUser.All");
    expect(provider!.scopes).toContain("https://outlook.office.com/SMTP.Send");
    expect(provider!.scopes).toContain("offline_access");
    expect(provider!.scopes).toContain("openid");
    expect(provider!.scopes).toContain("profile");
    expect(provider!.scopes).toContain("email");
    expect(provider!.userInfoUrl).toBeUndefined();
    expect(provider!.usePkce).toBe(true);
  });

  it("returns yahoo provider config", () => {
    const provider = getOAuthProvider("yahoo");
    expect(provider).not.toBeNull();
    expect(provider!.id).toBe("yahoo");
    expect(provider!.name).toBe("Yahoo");
    expect(provider!.authUrl).toContain("login.yahoo.com");
    expect(provider!.scopes).toContain("mail-r");
    expect(provider!.scopes).toContain("mail-w");
    expect(provider!.usePkce).toBe(true);
  });

  it("returns null for unknown provider", () => {
    expect(getOAuthProvider("unknown")).toBeNull();
  });

  it("returns null for empty string", () => {
    expect(getOAuthProvider("")).toBeNull();
  });
});

describe("getAllOAuthProviders", () => {
  it("returns all registered providers", () => {
    const providers = getAllOAuthProviders();
    expect(providers.length).toBeGreaterThanOrEqual(2);
    const ids = providers.map((p) => p.id);
    expect(ids).toContain("microsoft");
    expect(ids).toContain("yahoo");
  });
});

describe("getMicrosoftGraphProvider", () => {
  it("builds the authority from the tenant", () => {
    const provider = getMicrosoftGraphProvider("11111111-2222-3333-4444-555555555555");
    expect(provider.id).toBe("microsoft_graph");
    expect(provider.authUrl).toBe(
      "https://login.microsoftonline.com/11111111-2222-3333-4444-555555555555/oauth2/v2.0/authorize",
    );
    expect(provider.tokenUrl).toBe(
      "https://login.microsoftonline.com/11111111-2222-3333-4444-555555555555/oauth2/v2.0/token",
    );
    expect(provider.usePkce).toBe(true);
  });

  it("defaults to any work or school tenant", () => {
    expect(getMicrosoftGraphProvider().authUrl).toContain("/organizations/");
    expect(getMicrosoftGraphProvider("  ").authUrl).toContain("/organizations/");
  });

  it("accepts a domain as the tenant", () => {
    expect(getMicrosoftGraphProvider("contoso.com").tokenUrl).toContain("/contoso.com/");
  });

  it("rejects tenants that would change the URL", () => {
    expect(() => getMicrosoftGraphProvider("evil.com/x?y=")).toThrow(/Invalid Microsoft tenant/);
  });

  it("requests mail, calendar and offline scopes", () => {
    const { scopes } = getMicrosoftGraphProvider();
    expect(scopes).toContain("https://graph.microsoft.com/Mail.ReadWrite");
    expect(scopes).toContain("https://graph.microsoft.com/Mail.Send");
    expect(scopes).toContain("https://graph.microsoft.com/Calendars.ReadWrite");
    expect(scopes).toContain("offline_access");
  });

  it("is resolvable by ID and counted as Microsoft", () => {
    expect(getOAuthProvider("microsoft_graph")?.id).toBe("microsoft_graph");
    expect(isMicrosoftProvider("microsoft_graph")).toBe(true);
    expect(isMicrosoftProvider("microsoft")).toBe(true);
    expect(isMicrosoftProvider("yahoo")).toBe(false);
  });
});
