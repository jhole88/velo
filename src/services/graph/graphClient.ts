import { fetch } from "@tauri-apps/plugin-http";
import { getGraphAccessToken } from "./graphAuth";

export const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

/** Graph caps a JSON batch at 20 sub-requests. */
export const GRAPH_BATCH_LIMIT = 20;

const MAX_RETRIES = 3;
const MAX_RETRY_DELAY_MS = 30_000;

/**
 * Immutable IDs stay the same when a message is moved between folders, which
 * is what lets the local DB keep one row per message across archive/move.
 */
const IMMUTABLE_ID_PREFERENCE = 'IdType="ImmutableId"';

export class GraphError extends Error {
  constructor(
    readonly status: number,
    readonly code: string | null,
    message: string,
  ) {
    super(message);
    this.name = "GraphError";
  }
}

export interface GraphRequestOptions {
  method?: "GET" | "POST" | "PATCH" | "DELETE";
  /** JSON body (object) or a raw string body when `contentType` is set. */
  body?: unknown;
  contentType?: string;
  /** Extra `Prefer` header values, e.g. `odata.maxpagesize=100`. */
  prefer?: string[];
}

export interface GraphPage<T> {
  value: T[];
  "@odata.nextLink"?: string;
  "@odata.deltaLink"?: string;
}

export interface GraphBatchRequest {
  id: string;
  method: "GET" | "POST" | "PATCH" | "DELETE";
  /** Path relative to the API root, e.g. `/me/messages/{id}/move`. */
  url: string;
  body?: unknown;
}

export interface GraphBatchResponse {
  id: string;
  status: number;
  body?: unknown;
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function resolveUrl(pathOrUrl: string): string {
  return pathOrUrl.startsWith("https://") ? pathOrUrl : `${GRAPH_BASE_URL}${pathOrUrl}`;
}

function retryDelayMs(response: Response, attempt: number): number {
  const retryAfter = Number(response.headers.get("Retry-After"));
  const ms = Number.isFinite(retryAfter) && retryAfter > 0 ? retryAfter * 1000 : 1000 * 2 ** attempt;
  return Math.min(ms, MAX_RETRY_DELAY_MS);
}

async function toGraphError(response: Response): Promise<GraphError> {
  let code: string | null = null;
  let message = `Microsoft Graph request failed (${response.status})`;
  try {
    const data = (await response.json()) as { error?: { code?: string; message?: string } };
    if (data.error) {
      code = data.error.code ?? null;
      if (data.error.message) message = `${message}: ${data.error.message}`;
    }
  } catch {
    // Non-JSON error body — keep the generic message
  }
  return new GraphError(response.status, code, message);
}

/**
 * Thin Microsoft Graph client for one account: bearer auth with automatic
 * token refresh, throttling retries, paging and JSON batching.
 */
export class GraphClient {
  constructor(readonly accountId: string) {}

  private async send(pathOrUrl: string, options: GraphRequestOptions): Promise<Response> {
    const url = resolveUrl(pathOrUrl);
    let forceRefresh = false;

    for (let attempt = 0; ; attempt++) {
      const token = await getGraphAccessToken(this.accountId, forceRefresh);
      forceRefresh = false;

      const headers: Record<string, string> = {
        Authorization: `Bearer ${token}`,
        Prefer: [IMMUTABLE_ID_PREFERENCE, ...(options.prefer ?? [])].join(", "),
      };
      let body: string | undefined;
      if (options.body !== undefined) {
        if (options.contentType) {
          headers["Content-Type"] = options.contentType;
          body = String(options.body);
        } else {
          headers["Content-Type"] = "application/json";
          body = JSON.stringify(options.body);
        }
      }

      const response = await fetch(url, { method: options.method ?? "GET", headers, body });
      if (response.ok) return response;

      if (attempt < MAX_RETRIES) {
        if (response.status === 401) {
          // Token rejected — refresh once and retry
          forceRefresh = true;
          continue;
        }
        if (response.status === 429 || response.status === 503 || response.status === 504) {
          await delay(retryDelayMs(response, attempt));
          continue;
        }
      }

      throw await toGraphError(response);
    }
  }

  /** Send a request and parse the JSON response (undefined for empty bodies). */
  async request<T>(pathOrUrl: string, options: GraphRequestOptions = {}): Promise<T> {
    const response = await this.send(pathOrUrl, options);
    if (response.status === 202 || response.status === 204) return undefined as T;
    const text = await response.text();
    return (text ? JSON.parse(text) : undefined) as T;
  }

  /** Send a GET and return the raw response body as text (e.g. MIME content). */
  async requestText(pathOrUrl: string): Promise<string> {
    const response = await this.send(pathOrUrl, {});
    return response.text();
  }

  /** Send a GET and return the raw response body as bytes (e.g. attachments). */
  async requestBytes(pathOrUrl: string): Promise<Uint8Array> {
    const response = await this.send(pathOrUrl, {});
    return new Uint8Array(await response.arrayBuffer());
  }

  /**
   * Follow `@odata.nextLink` until the collection is exhausted, handing each
   * page to `onPage`. Returns the final `@odata.deltaLink` when present.
   */
  async forEachPage<T>(
    pathOrUrl: string,
    onPage: (items: T[]) => Promise<void> | void,
    options: GraphRequestOptions = {},
  ): Promise<string | null> {
    let next: string | undefined = pathOrUrl;
    let deltaLink: string | null = null;

    while (next) {
      const page: GraphPage<T> = await this.request<GraphPage<T>>(next, options);
      await onPage(page.value ?? []);
      deltaLink = page["@odata.deltaLink"] ?? deltaLink;
      next = page["@odata.nextLink"];
    }

    return deltaLink;
  }

  /** Collect every item of a paged collection. */
  async getAll<T>(pathOrUrl: string, options: GraphRequestOptions = {}): Promise<T[]> {
    const all: T[] = [];
    await this.forEachPage<T>(pathOrUrl, (items) => { all.push(...items); }, options);
    return all;
  }

  /**
   * Run requests through the JSON batch endpoint in chunks of 20.
   * Sub-requests that fail are returned with their status — callers decide
   * what is fatal (see `assertBatchOk`).
   */
  async batch(requests: GraphBatchRequest[]): Promise<GraphBatchResponse[]> {
    const results: GraphBatchResponse[] = [];

    for (let i = 0; i < requests.length; i += GRAPH_BATCH_LIMIT) {
      const chunk = requests.slice(i, i + GRAPH_BATCH_LIMIT);
      const payload = {
        requests: chunk.map((req) => ({
          id: req.id,
          method: req.method,
          url: req.url,
          headers: {
            Prefer: IMMUTABLE_ID_PREFERENCE,
            ...(req.body !== undefined ? { "Content-Type": "application/json" } : {}),
          },
          ...(req.body !== undefined ? { body: req.body } : {}),
        })),
      };
      const response = await this.request<{ responses: GraphBatchResponse[] }>("/$batch", {
        method: "POST",
        body: payload,
      });
      results.push(...(response?.responses ?? []));
    }

    return results;
  }
}

/**
 * Throw if any batch sub-request failed. 404s are tolerated by default: the
 * item is already gone on the server, which is the state the caller wanted.
 */
export function assertBatchOk(responses: GraphBatchResponse[], tolerateNotFound = true): void {
  for (const res of responses) {
    if (res.status < 400) continue;
    if (tolerateNotFound && res.status === 404) continue;
    const error = (res.body as { error?: { code?: string; message?: string } } | undefined)?.error;
    throw new GraphError(
      res.status,
      error?.code ?? null,
      `Microsoft Graph batch request failed (${res.status})${error?.message ? `: ${error.message}` : ""}`,
    );
  }
}

const clients = new Map<string, GraphClient>();

export function getGraphClient(accountId: string): GraphClient {
  let client = clients.get(accountId);
  if (!client) {
    client = new GraphClient(accountId);
    clients.set(accountId, client);
  }
  return client;
}

export function removeGraphClient(accountId: string): void {
  clients.delete(accountId);
}

/**
 * Look up the signed-in mailbox with a bare access token (used during account
 * setup, before an account row exists). The ID token's username can differ
 * from the primary mail address, so Graph is the source of truth.
 */
export async function fetchGraphProfile(
  accessToken: string,
): Promise<{ email: string | null; name: string | null }> {
  const response = await fetch(`${GRAPH_BASE_URL}/me?$select=mail,userPrincipalName,displayName`, {
    method: "GET",
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!response.ok) throw await toGraphError(response);
  const me = (await response.json()) as {
    mail?: string | null;
    userPrincipalName?: string | null;
    displayName?: string | null;
  };
  return { email: me.mail ?? me.userPrincipalName ?? null, name: me.displayName ?? null };
}
