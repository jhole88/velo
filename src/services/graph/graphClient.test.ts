import { describe, it, expect, beforeEach, vi } from "vitest";

vi.mock("@tauri-apps/plugin-http", () => ({ fetch: vi.fn() }));
vi.mock("./graphAuth", () => ({ getGraphAccessToken: vi.fn() }));

import { fetch } from "@tauri-apps/plugin-http";
import { getGraphAccessToken } from "./graphAuth";
import { GraphClient, GraphError, assertBatchOk } from "./graphClient";

function jsonResponse(status: number, body: unknown, headers: Record<string, string> = {}): Response {
  return new Response(body === undefined ? null : JSON.stringify(body), { status, headers });
}

describe("GraphClient", () => {
  let client: GraphClient;

  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(getGraphAccessToken).mockResolvedValue("token-1");
    client = new GraphClient("acc-1");
  });

  it("sends bearer auth and the immutable ID preference", async () => {
    vi.mocked(fetch).mockResolvedValue(jsonResponse(200, { id: "x" }));

    const result = await client.request<{ id: string }>("/me", { prefer: ["odata.maxpagesize=10"] });

    expect(result).toEqual({ id: "x" });
    const [url, init] = vi.mocked(fetch).mock.calls[0]!;
    expect(url).toBe("https://graph.microsoft.com/v1.0/me");
    const headers = (init as { headers: Record<string, string> }).headers;
    expect(headers.Authorization).toBe("Bearer token-1");
    expect(headers.Prefer).toBe('IdType="ImmutableId", odata.maxpagesize=10');
  });

  it("passes absolute URLs through untouched", async () => {
    vi.mocked(fetch).mockResolvedValue(jsonResponse(200, { value: [] }));
    await client.request("https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc");
    expect(vi.mocked(fetch).mock.calls[0]![0]).toBe("https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc");
  });

  it("sends raw bodies with the given content type", async () => {
    vi.mocked(fetch).mockResolvedValue(new Response(null, { status: 202 }));
    const result = await client.request("/me/sendMail", { method: "POST", contentType: "text/plain", body: "TUlNRQ==" });
    expect(result).toBeUndefined();
    const init = vi.mocked(fetch).mock.calls[0]![1] as { headers: Record<string, string>; body: string };
    expect(init.headers["Content-Type"]).toBe("text/plain");
    expect(init.body).toBe("TUlNRQ==");
  });

  it("refreshes the token and retries after a 401", async () => {
    vi.mocked(fetch)
      .mockResolvedValueOnce(jsonResponse(401, { error: { code: "InvalidAuthenticationToken" } }))
      .mockResolvedValueOnce(jsonResponse(200, { ok: true }));

    await client.request("/me");

    expect(getGraphAccessToken).toHaveBeenNthCalledWith(1, "acc-1", false);
    expect(getGraphAccessToken).toHaveBeenNthCalledWith(2, "acc-1", true);
  });

  it("waits and retries when throttled", async () => {
    vi.mocked(fetch)
      .mockResolvedValueOnce(jsonResponse(429, { error: { code: "TooManyRequests" } }, { "Retry-After": "0.01" }))
      .mockResolvedValueOnce(jsonResponse(200, { ok: true }));

    await expect(client.request("/me")).resolves.toEqual({ ok: true });
    expect(fetch).toHaveBeenCalledTimes(2);
  });

  it("throws a GraphError with status and code", async () => {
    vi.mocked(fetch).mockResolvedValue(jsonResponse(404, { error: { code: "ErrorItemNotFound", message: "Not found" } }));

    const err = await client.request("/me/messages/x").catch((e: unknown) => e);
    expect(err).toBeInstanceOf(GraphError);
    expect((err as GraphError).status).toBe(404);
    expect((err as GraphError).code).toBe("ErrorItemNotFound");
  });

  it("follows nextLink and returns the delta link", async () => {
    vi.mocked(fetch)
      .mockResolvedValueOnce(jsonResponse(200, { value: [{ id: "1" }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next" }))
      .mockResolvedValueOnce(jsonResponse(200, { value: [{ id: "2" }], "@odata.deltaLink": "https://graph.microsoft.com/v1.0/delta" }));

    const seen: string[] = [];
    const deltaLink = await client.forEachPage<{ id: string }>("/me/mailFolders/inbox/messages/delta", (items) => {
      seen.push(...items.map((i) => i.id));
    });

    expect(seen).toEqual(["1", "2"]);
    expect(deltaLink).toBe("https://graph.microsoft.com/v1.0/delta");
  });

  it("splits batches into chunks of 20", async () => {
    vi.mocked(fetch).mockImplementation(async (_url, init) => {
      const payload = JSON.parse((init as { body: string }).body) as { requests: { id: string }[] };
      return jsonResponse(200, { responses: payload.requests.map((r) => ({ id: r.id, status: 204 })) });
    });

    const requests = Array.from({ length: 45 }, (_, i) => ({
      id: String(i),
      method: "PATCH" as const,
      url: `/me/messages/${i}`,
      body: { isRead: true },
    }));
    const responses = await client.batch(requests);

    expect(fetch).toHaveBeenCalledTimes(3);
    expect(responses).toHaveLength(45);
    const first = JSON.parse((vi.mocked(fetch).mock.calls[0]![1] as { body: string }).body) as {
      requests: { headers: Record<string, string>; body: unknown }[];
    };
    expect(first.requests).toHaveLength(20);
    expect(first.requests[0]!.headers).toEqual({ Prefer: 'IdType="ImmutableId"', "Content-Type": "application/json" });
  });
});

describe("assertBatchOk", () => {
  it("tolerates 404s by default", () => {
    expect(() => assertBatchOk([{ id: "0", status: 204 }, { id: "1", status: 404 }])).not.toThrow();
  });

  it("throws on other failures", () => {
    expect(() =>
      assertBatchOk([{ id: "0", status: 403, body: { error: { code: "ErrorAccessDenied", message: "Denied" } } }]),
    ).toThrow(GraphError);
  });

  it("throws on 404 when not tolerated", () => {
    expect(() => assertBatchOk([{ id: "0", status: 404 }], false)).toThrow(GraphError);
  });
});
