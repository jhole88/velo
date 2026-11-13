import { describe, it, expect, beforeEach, vi } from "vitest";

const mockDb = { select: vi.fn(), execute: vi.fn() };

vi.mock("../db/connection", () => ({
  getDb: vi.fn(async () => mockDb),
  withTransaction: vi.fn(async (fn: (db: unknown) => Promise<void>) => fn(mockDb)),
}));
vi.mock("../db/accounts", () => ({ updateAccountSyncState: vi.fn() }));
vi.mock("../db/labels", () => ({
  upsertLabel: vi.fn(),
  getLabelsForAccount: vi.fn(async () => []),
  deleteLabel: vi.fn(),
}));
vi.mock("../db/threads", () => ({
  upsertThread: vi.fn(),
  setThreadLabels: vi.fn(),
  deleteThread: vi.fn(),
  getMutedThreadIds: vi.fn(async () => new Set<string>()),
}));
vi.mock("../db/messages", () => ({ upsertMessage: vi.fn(), deleteMessage: vi.fn() }));
vi.mock("../db/attachments", () => ({ upsertAttachment: vi.fn() }));
vi.mock("../db/settings", () => ({ getSetting: vi.fn(async () => null), setSetting: vi.fn() }));
vi.mock("../db/pendingOperations", () => ({ getPendingOpsForResource: vi.fn(async () => []) }));
vi.mock("../db/threadCategories", () => ({
  getThreadCategory: vi.fn(async () => "Primary"),
  getThreadCategoryWithManual: vi.fn(async () => null),
  setThreadCategory: vi.fn(),
}));
vi.mock("../db/notificationVips", () => ({ getVipSenders: vi.fn(async () => new Set<string>()) }));
vi.mock("../db/bundleRules", () => ({
  getBundleRule: vi.fn(async () => null),
  holdThread: vi.fn(),
  getNextDeliveryTime: vi.fn(),
}));
vi.mock("../notifications/notificationManager", () => ({
  shouldNotifyForMessage: vi.fn(() => true),
  queueNewEmailNotification: vi.fn(),
}));
vi.mock("../filters/filterEngine", () => ({ applyFiltersToMessages: vi.fn() }));
vi.mock("@/services/smartLabels/smartLabelManager", () => ({ applySmartLabelsToMessages: vi.fn(async () => undefined) }));
vi.mock("@/services/ai/categorizationManager", () => ({ categorizeNewThreads: vi.fn(async () => undefined) }));
vi.mock("./graphClient", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./graphClient")>();
  return { ...actual, getGraphClient: vi.fn() };
});

import { updateAccountSyncState } from "../db/accounts";
import { upsertLabel, getLabelsForAccount, deleteLabel } from "../db/labels";
import { upsertThread, setThreadLabels, deleteThread } from "../db/threads";
import { upsertMessage, deleteMessage } from "../db/messages";
import { upsertAttachment } from "../db/attachments";
import { getSetting, setSetting } from "../db/settings";
import { getPendingOpsForResource } from "../db/pendingOperations";
import { queueNewEmailNotification } from "../notifications/notificationManager";
import { applyFiltersToMessages } from "../filters/filterEngine";
import { GraphError, getGraphClient, type GraphBatchRequest, type GraphClient } from "./graphClient";
import { graphDeltaSync, graphInitialSync, resetGraphSyncFlags } from "./graphSync";
import type { GraphMessage } from "./graphMapper";

const INBOX_DELTA = "https://graph.microsoft.com/v1.0/delta-inbox";
const SENT_DELTA = "https://graph.microsoft.com/v1.0/delta-sent";

function message(overrides: Partial<GraphMessage>): GraphMessage {
  return {
    id: "m1",
    conversationId: "conv-1",
    parentFolderId: "f-inbox",
    subject: "Hello",
    bodyPreview: "preview",
    body: { contentType: "html", content: "<p>hi</p>" },
    from: { emailAddress: { name: "Sam", address: "sam@example.com" } },
    receivedDateTime: "2026-11-01T12:00:00Z",
    isRead: true,
    ...overrides,
  };
}

interface FakeMailbox {
  inboxMessages: GraphMessage[];
  sentMessages: GraphMessage[];
  conversations: Record<string, GraphMessage[]>;
  deltaItems: Record<string, GraphMessage[]>;
  deltaErrors: Record<string, Error>;
  rejectContentIdSelect: boolean;
}

/** Minimal stand-in for GraphClient that answers by URL. */
function createFakeClient(mailbox: FakeMailbox) {
  const calls: string[] = [];

  const pages = (url: string): { items: unknown[]; deltaLink: string | null } => {
    calls.push(url);
    if (url.startsWith("/me/mailFolders?")) {
      return {
        items: [
          { id: "f-inbox", displayName: "Inbox", parentFolderId: "root", totalItemCount: 2, childFolderCount: 0 },
          { id: "f-sent", displayName: "Sent Items", parentFolderId: "root", totalItemCount: 1, childFolderCount: 0 },
          { id: "f-outbox", displayName: "Outbox", parentFolderId: "root", totalItemCount: 0, childFolderCount: 0 },
        ],
        deltaLink: null,
      };
    }
    if (url === "/me/outlook/masterCategories") {
      return { items: [{ displayName: "Blue category", color: "preset7" }], deltaLink: null };
    }
    if (url.includes("/messages/delta")) {
      return { items: [], deltaLink: url.includes("f-inbox") ? INBOX_DELTA : SENT_DELTA };
    }
    if (url === INBOX_DELTA || url === SENT_DELTA) {
      const err = mailbox.deltaErrors[url];
      if (err) throw err;
      return { items: mailbox.deltaItems[url] ?? [], deltaLink: `${url}-next` };
    }
    if (url.startsWith("/me/mailFolders/f-inbox/messages?")) return { items: mailbox.inboxMessages, deltaLink: null };
    if (url.startsWith("/me/mailFolders/f-sent/messages?")) return { items: mailbox.sentMessages, deltaLink: null };
    if (url.startsWith("/me/messages?")) {
      if (mailbox.rejectContentIdSelect && url.includes("fileAttachment")) {
        throw new GraphError(400, "BadRequest", "Could not find a property named 'contentId'");
      }
      const match = /conversationId%20eq%20'([^']+)'/.exec(url);
      return { items: mailbox.conversations[decodeURIComponent(match?.[1] ?? "")] ?? [], deltaLink: null };
    }
    throw new Error(`Unexpected Graph URL in test: ${url}`);
  };

  const client = {
    accountId: "acc-1",
    forEachPage: vi.fn(async (url: string, onPage: (items: unknown[]) => Promise<void> | void) => {
      const { items, deltaLink } = pages(url);
      await onPage(items);
      return deltaLink;
    }),
    getAll: vi.fn(async (url: string) => pages(url).items),
    batch: vi.fn(async (requests: GraphBatchRequest[]) =>
      requests.map((req) => {
        if (req.method === "GET" && req.url.startsWith("/me/mailFolders/")) {
          const ids: Record<string, string> = { inbox: "f-inbox", sentitems: "f-sent", outbox: "f-outbox" };
          const id = ids[req.id];
          return id ? { id: req.id, status: 200, body: { id } } : { id: req.id, status: 404 };
        }
        return { id: req.id, status: 201 };
      }),
    ),
    request: vi.fn(),
  };

  return { client: client as unknown as GraphClient, calls, batch: client.batch };
}

function emptyMailbox(): FakeMailbox {
  return {
    inboxMessages: [],
    sentMessages: [],
    conversations: {},
    deltaItems: {},
    deltaErrors: {},
    rejectContentIdSelect: false,
  };
}

function storedState(): string {
  return JSON.stringify({ deltaLinks: { "f-inbox": INBOX_DELTA, "f-sent": SENT_DELTA }, retryThreadIds: [] });
}

describe("graphSync", () => {
  let mailbox: FakeMailbox;

  beforeEach(() => {
    vi.clearAllMocks();
    resetGraphSyncFlags();
    mailbox = emptyMailbox();
    mockDb.select.mockResolvedValue([]);
    mockDb.execute.mockResolvedValue(undefined);
    vi.mocked(getSetting).mockResolvedValue(null);
    vi.mocked(getLabelsForAccount).mockResolvedValue([]);
    vi.mocked(getPendingOpsForResource).mockResolvedValue([]);
  });

  describe("graphInitialSync", () => {
    it("mirrors folders and categories as labels and skips Outbox", async () => {
      const { client } = createFakeClient(mailbox);
      vi.mocked(getGraphClient).mockReturnValue(client);

      await graphInitialSync("acc-1", 30);

      const labelIds = vi.mocked(upsertLabel).mock.calls.map(([label]) => label.id);
      expect(labelIds).toEqual(["INBOX", "SENT", "graph-cat:Blue category"]);
      expect(upsertLabel).toHaveBeenCalledWith(
        expect.objectContaining({ id: "INBOX", type: "system", imapFolderPath: "f-inbox" }),
      );
      expect(upsertLabel).toHaveBeenCalledWith(
        expect.objectContaining({ id: "graph-cat:Blue category", colorBg: "#00bcf2" }),
      );
    });

    it("removes labels for folders that no longer exist", async () => {
      const { client } = createFakeClient(mailbox);
      vi.mocked(getGraphClient).mockReturnValue(client);
      vi.mocked(getLabelsForAccount).mockResolvedValue([
        { id: "INBOX" }, { id: "graph-folder:gone" }, { id: "graph-cat:Old" },
      ] as never);

      await graphInitialSync("acc-1", 30);

      expect(deleteLabel).toHaveBeenCalledWith("acc-1", "graph-folder:gone");
      expect(deleteLabel).toHaveBeenCalledWith("acc-1", "graph-cat:Old");
      expect(deleteLabel).not.toHaveBeenCalledWith("acc-1", "INBOX");
    });

    it("stores messages and combines a conversation spread across folders into one thread", async () => {
      mailbox.inboxMessages = [
        message({ id: "m2", subject: "RE: Hello", bodyPreview: "latest", receivedDateTime: "2026-11-02T12:00:00Z", isRead: false,
          attachments: [{ id: "att-1", name: "a.pdf", contentType: "application/pdf", size: 10, isInline: false }] }),
      ];
      mailbox.sentMessages = [
        message({ id: "m1", parentFolderId: "f-sent", subject: "Hello", bodyPreview: "first", receivedDateTime: "2026-11-01T12:00:00Z" }),
      ];
      const { client } = createFakeClient(mailbox);
      vi.mocked(getGraphClient).mockReturnValue(client);

      await graphInitialSync("acc-1", 30);

      expect(upsertMessage).toHaveBeenCalledTimes(2);
      expect(upsertMessage).toHaveBeenCalledWith(expect.objectContaining({ id: "m2", threadId: "conv-1", imapFolder: "f-inbox" }));
      expect(upsertMessage).toHaveBeenCalledWith(expect.objectContaining({ id: "m1", threadId: "conv-1", imapFolder: "f-sent" }));
      expect(upsertAttachment).toHaveBeenCalledWith(expect.objectContaining({ id: "m2_att-1", gmailAttachmentId: "att-1" }));

      const lastThread = vi.mocked(upsertThread).mock.calls.at(-1)![0];
      expect(lastThread).toMatchObject({
        id: "conv-1",
        subject: "Hello", // earliest message
        snippet: "latest", // newest message
        messageCount: 2,
        isRead: false,
        hasAttachments: true,
      });
      const lastLabels = vi.mocked(setThreadLabels).mock.calls.at(-1)!;
      expect(lastLabels[2]).toEqual(expect.arrayContaining(["INBOX", "SENT", "UNREAD"]));
    });

    it("saves a delta link per folder and marks the account as synced", async () => {
      const { client } = createFakeClient(mailbox);
      vi.mocked(getGraphClient).mockReturnValue(client);

      await graphInitialSync("acc-1", 30);

      const [key, value] = vi.mocked(setSetting).mock.calls.at(-1)!;
      expect(key).toBe("graph_sync_state:acc-1");
      expect(JSON.parse(value)).toEqual({
        deltaLinks: { "f-inbox": INBOX_DELTA, "f-sent": SENT_DELTA },
        retryThreadIds: [],
      });
      expect(updateAccountSyncState).toHaveBeenCalledWith("acc-1", expect.stringMatching(/^graph:\d+$/));
    });

    it("reports progress through to done", async () => {
      mailbox.inboxMessages = [message({})];
      const { client } = createFakeClient(mailbox);
      vi.mocked(getGraphClient).mockReturnValue(client);
      const phases: string[] = [];

      await graphInitialSync("acc-1", 30, (p) => phases.push(p.phase));

      expect(phases[0]).toBe("labels");
      expect(phases).toContain("messages");
      expect(phases.at(-1)).toBe("done");
    });
  });

  describe("graphDeltaSync", () => {
    beforeEach(() => {
      vi.mocked(getSetting).mockImplementation(async (key: string) =>
        key === "graph_sync_state:acc-1" ? storedState() : null,
      );
    });

    it("does nothing but advance the delta links when nothing changed", async () => {
      const { client, calls } = createFakeClient(mailbox);
      vi.mocked(getGraphClient).mockReturnValue(client);

      const refreshed = await graphDeltaSync("acc-1");

      expect(refreshed).toBe(0);
      expect(upsertThread).not.toHaveBeenCalled();
      expect(calls.some((url) => url.startsWith("/me/messages?"))).toBe(false);
      const saved = JSON.parse(vi.mocked(setSetting).mock.calls.at(-1)![1]);
      expect(saved.deltaLinks).toEqual({ "f-inbox": `${INBOX_DELTA}-next`, "f-sent": `${SENT_DELTA}-next` });
    });

    it("re-downloads a changed conversation, notifies and applies filters for new unread inbox mail", async () => {
      const incoming = message({ id: "m-new", conversationId: "conv-9", isRead: false, subject: "New RFI" });
      mailbox.deltaItems[INBOX_DELTA] = [{ id: "m-new", conversationId: "conv-9", isRead: false }];
      mailbox.conversations["conv-9"] = [incoming];
      const { client } = createFakeClient(mailbox);
      vi.mocked(getGraphClient).mockReturnValue(client);

      const refreshed = await graphDeltaSync("acc-1");

      expect(refreshed).toBe(1);
      expect(upsertThread).toHaveBeenCalledWith(expect.objectContaining({ id: "conv-9", isRead: false }));
      expect(upsertMessage).toHaveBeenCalledWith(expect.objectContaining({ id: "m-new", threadId: "conv-9" }));
      expect(queueNewEmailNotification).toHaveBeenCalledWith("Sam", "New RFI", "conv-9", "acc-1", "sam@example.com");
      expect(applyFiltersToMessages).toHaveBeenCalledWith("acc-1", [expect.objectContaining({ id: "m-new" })]);
    });

    it("does not notify for messages already stored locally", async () => {
      mailbox.deltaItems[INBOX_DELTA] = [{ id: "m-old", conversationId: "conv-9", isRead: false }];
      mailbox.conversations["conv-9"] = [message({ id: "m-old", conversationId: "conv-9", isRead: false })];
      mockDb.select.mockImplementation(async (sql: string) =>
        sql.includes("SELECT id FROM messages") ? [{ id: "m-old" }] : [],
      );
      const { client } = createFakeClient(mailbox);
      vi.mocked(getGraphClient).mockReturnValue(client);

      await graphDeltaSync("acc-1");

      expect(queueNewEmailNotification).not.toHaveBeenCalled();
      expect(applyFiltersToMessages).not.toHaveBeenCalled();
    });

    it("deletes the local thread when its last message is gone from the server", async () => {
      mailbox.deltaItems[INBOX_DELTA] = [{ id: "m-del", "@removed": { reason: "deleted" } }];
      mailbox.conversations["conv-del"] = [];
      mockDb.select.mockImplementation(async (sql: string) => {
        if (sql.includes("SELECT DISTINCT thread_id")) return [{ thread_id: "conv-del" }];
        if (sql.includes("SELECT id FROM messages")) return [{ id: "m-del" }];
        return [];
      });
      const { client } = createFakeClient(mailbox);
      vi.mocked(getGraphClient).mockReturnValue(client);

      await graphDeltaSync("acc-1");

      expect(deleteMessage).toHaveBeenCalledWith("acc-1", "m-del");
      expect(deleteThread).toHaveBeenCalledWith("acc-1", "conv-del");
    });

    it("drops local messages that left a thread but keeps the thread", async () => {
      mailbox.deltaItems[INBOX_DELTA] = [{ id: "m-gone", "@removed": { reason: "deleted" } }];
      mailbox.conversations["conv-1"] = [message({ id: "m-stay" })];
      mockDb.select.mockImplementation(async (sql: string) => {
        if (sql.includes("SELECT DISTINCT thread_id")) return [{ thread_id: "conv-1" }];
        if (sql.includes("SELECT id FROM messages")) return [{ id: "m-gone" }, { id: "m-stay" }];
        return [];
      });
      const { client } = createFakeClient(mailbox);
      vi.mocked(getGraphClient).mockReturnValue(client);

      await graphDeltaSync("acc-1");

      expect(deleteMessage).toHaveBeenCalledWith("acc-1", "m-gone");
      expect(deleteMessage).not.toHaveBeenCalledWith("acc-1", "m-stay");
      expect(deleteThread).not.toHaveBeenCalled();
      expect(upsertThread).toHaveBeenCalledWith(expect.objectContaining({ id: "conv-1", messageCount: 1 }));
    });

    it("ignores messages in folders that are not synced", async () => {
      mailbox.deltaItems[INBOX_DELTA] = [{ id: "m1", conversationId: "conv-1" }];
      mailbox.conversations["conv-1"] = [
        message({ id: "m1" }),
        message({ id: "m-outbox", parentFolderId: "f-outbox" }),
      ];
      const { client } = createFakeClient(mailbox);
      vi.mocked(getGraphClient).mockReturnValue(client);

      await graphDeltaSync("acc-1");

      expect(upsertMessage).toHaveBeenCalledTimes(1);
      expect(upsertMessage).toHaveBeenCalledWith(expect.objectContaining({ id: "m1" }));
    });

    it("leaves threads with pending local changes alone", async () => {
      mailbox.deltaItems[INBOX_DELTA] = [{ id: "m1", conversationId: "conv-1" }];
      mailbox.conversations["conv-1"] = [message({})];
      vi.mocked(getPendingOpsForResource).mockResolvedValue([{ id: "op" }] as never);
      const { client } = createFakeClient(mailbox);
      vi.mocked(getGraphClient).mockReturnValue(client);

      await graphDeltaSync("acc-1");

      expect(upsertThread).not.toHaveBeenCalled();
    });

    it("rebuilds a folder's baseline when its delta link has expired", async () => {
      mailbox.deltaErrors[INBOX_DELTA] = new GraphError(410, "SyncStateNotFound", "expired");
      mailbox.inboxMessages = [{ id: "m1", conversationId: "conv-1" }];
      mailbox.conversations["conv-1"] = [message({})];
      const { client } = createFakeClient(mailbox);
      vi.mocked(getGraphClient).mockReturnValue(client);

      await graphDeltaSync("acc-1");

      const saved = JSON.parse(vi.mocked(setSetting).mock.calls.at(-1)![1]);
      expect(saved.deltaLinks["f-inbox"]).toBe(INBOX_DELTA); // fresh baseline from the fake
      expect(upsertThread).toHaveBeenCalledWith(expect.objectContaining({ id: "conv-1" }));
    });

    it("remembers conversations whose refresh failed so the next sync retries them", async () => {
      mailbox.deltaItems[INBOX_DELTA] = [{ id: "m1", conversationId: "conv-1" }];
      mailbox.conversations["conv-1"] = [message({})];
      vi.mocked(upsertThread).mockRejectedValueOnce(new Error("db locked"));
      const { client } = createFakeClient(mailbox);
      vi.mocked(getGraphClient).mockReturnValue(client);

      await graphDeltaSync("acc-1");

      const saved = JSON.parse(vi.mocked(setSetting).mock.calls.at(-1)![1]);
      expect(saved.retryThreadIds).toEqual(["conv-1"]);
    });

    it("falls back to the basic attachment select when the contentId cast is rejected", async () => {
      mailbox.rejectContentIdSelect = true;
      mailbox.deltaItems[INBOX_DELTA] = [{ id: "m1", conversationId: "conv-1" }];
      mailbox.conversations["conv-1"] = [message({})];
      const { client, calls } = createFakeClient(mailbox);
      vi.mocked(getGraphClient).mockReturnValue(client);

      await graphDeltaSync("acc-1");

      const conversationCalls = calls.filter((url) => url.startsWith("/me/messages?"));
      expect(conversationCalls).toHaveLength(2);
      expect(conversationCalls[1]).not.toContain("fileAttachment");
      expect(upsertMessage).toHaveBeenCalledWith(expect.objectContaining({ id: "m1" }));
    });
  });
});
