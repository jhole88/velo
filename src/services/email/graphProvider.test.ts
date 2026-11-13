import { describe, it, expect, beforeEach, vi } from "vitest";

const mockDb = { select: vi.fn(), execute: vi.fn() };

vi.mock("../db/connection", () => ({ getDb: vi.fn(async () => mockDb) }));
vi.mock("../db/labels", () => ({ getLabelsForAccount: vi.fn() }));
vi.mock("../graph/graphSync", () => ({
  graphInitialSync: vi.fn(),
  graphDeltaSync: vi.fn(),
  loadFolderContext: vi.fn(),
  moveMessages: vi.fn(),
}));
vi.mock("../gmail/syncManager", () => ({ syncAccount: vi.fn(async () => undefined) }));

import { getLabelsForAccount } from "../db/labels";
import { graphDeltaSync, graphInitialSync, loadFolderContext, moveMessages } from "../graph/graphSync";
import type { GraphBatchRequest, GraphClient } from "../graph/graphClient";
import { GraphProvider } from "./graphProvider";

const LABELS = [
  { id: "INBOX", imap_folder_path: "f-inbox" },
  { id: "SENT", imap_folder_path: "f-sent" },
  { id: "TRASH", imap_folder_path: "f-trash" },
  { id: "SPAM", imap_folder_path: "f-junk" },
  { id: "archive", imap_folder_path: "f-archive" },
  { id: "graph-folder:f-jobs", imap_folder_path: "f-jobs" },
  { id: "graph-cat:Blue category", imap_folder_path: null },
];

/** Thread with one inbox message, one sent reply and one archived message. */
const THREAD_ROWS = [
  { id: "m-inbox", imap_folder: "f-inbox", date: 300 },
  { id: "m-sent", imap_folder: "f-sent", date: 200 },
  { id: "m-archived", imap_folder: "f-archive", date: 100 },
];

function createMockClient() {
  return {
    accountId: "acc-1",
    request: vi.fn(),
    requestText: vi.fn(),
    requestBytes: vi.fn(),
    getAll: vi.fn(),
    forEachPage: vi.fn(),
    batch: vi.fn(async (requests: GraphBatchRequest[]) => requests.map((r) => ({ id: r.id, status: 200, body: {} }))),
  };
}

describe("GraphProvider", () => {
  let client: ReturnType<typeof createMockClient>;
  let provider: GraphProvider;

  beforeEach(() => {
    vi.clearAllMocks();
    client = createMockClient();
    provider = new GraphProvider("acc-1", client as unknown as GraphClient);
    mockDb.select.mockResolvedValue(THREAD_ROWS);
    mockDb.execute.mockResolvedValue(undefined);
    vi.mocked(getLabelsForAccount).mockResolvedValue(LABELS as never);
  });

  it("has correct accountId and type", () => {
    expect(provider.accountId).toBe("acc-1");
    expect(provider.type).toBe("graph");
  });

  describe("thread actions", () => {
    it("archive moves only the inbox messages and records their new folder", async () => {
      await provider.archive("conv-1", []);

      expect(moveMessages).toHaveBeenCalledWith(client, ["m-inbox"], "archive");
      const [sql, params] = mockDb.execute.mock.calls[0]!;
      expect(sql).toContain("UPDATE messages SET imap_folder");
      expect(params).toEqual(["f-archive", "acc-1", "m-inbox"]);
    });

    it("archive does nothing when no message is in the inbox", async () => {
      mockDb.select.mockResolvedValue([{ id: "m-sent", imap_folder: "f-sent", date: 1 }]);
      await provider.archive("conv-1", []);
      expect(moveMessages).not.toHaveBeenCalled();
    });

    it("trash moves everything that is not already in Deleted Items", async () => {
      mockDb.select.mockResolvedValue([...THREAD_ROWS, { id: "m-trashed", imap_folder: "f-trash", date: 50 }]);
      await provider.trash("conv-1", []);
      expect(moveMessages).toHaveBeenCalledWith(client, ["m-inbox", "m-sent", "m-archived"], "deleteditems");
    });

    it("limits an action to the given message IDs", async () => {
      await provider.trash("conv-1", ["m-sent"]);
      expect(moveMessages).toHaveBeenCalledWith(client, ["m-sent"], "deleteditems");
    });

    it("permanentDelete deletes each message", async () => {
      await provider.permanentDelete("conv-1", []);
      const requests = client.batch.mock.calls[0]![0];
      expect(requests.map((r) => [r.method, r.url])).toEqual([
        ["DELETE", "/me/messages/m-inbox"],
        ["DELETE", "/me/messages/m-sent"],
        ["DELETE", "/me/messages/m-archived"],
      ]);
    });

    it("markRead patches every message", async () => {
      await provider.markRead("conv-1", [], true);
      const requests = client.batch.mock.calls[0]![0];
      expect(requests).toHaveLength(3);
      expect(requests[0]).toMatchObject({ method: "PATCH", url: "/me/messages/m-inbox", body: { isRead: true } });
    });

    it("star flags only the latest message", async () => {
      await provider.star("conv-1", [], true);
      const requests = client.batch.mock.calls[0]![0];
      expect(requests).toHaveLength(1);
      expect(requests[0]).toMatchObject({ url: "/me/messages/m-inbox", body: { flag: { flagStatus: "flagged" } } });
    });

    it("unstar clears the flag on every message", async () => {
      await provider.star("conv-1", [], false);
      const requests = client.batch.mock.calls[0]![0];
      expect(requests).toHaveLength(3);
      expect(requests[0]!.body).toEqual({ flag: { flagStatus: "notFlagged" } });
    });

    it("spam moves received mail to Junk and leaves sent copies", async () => {
      await provider.spam("conv-1", [], true);
      expect(moveMessages).toHaveBeenCalledWith(client, ["m-inbox", "m-archived"], "junkemail");
    });

    it("not-spam moves junk mail back to the inbox", async () => {
      mockDb.select.mockResolvedValue([{ id: "m-junk", imap_folder: "f-junk", date: 1 }]);
      await provider.spam("conv-1", [], false);
      expect(moveMessages).toHaveBeenCalledWith(client, ["m-junk"], "inbox");
    });

    it("moveToFolder resolves folder labels and skips sent copies", async () => {
      await provider.moveToFolder("conv-1", [], "graph-folder:f-jobs");
      expect(moveMessages).toHaveBeenCalledWith(client, ["m-inbox", "m-archived"], "f-jobs");
    });

    it("fails when a batch sub-request is rejected", async () => {
      client.batch.mockResolvedValue([{ id: "0", status: 403, body: { error: { message: "Denied" } } }]);
      await expect(provider.markRead("conv-1", ["m-inbox"], true)).rejects.toThrow(/Denied/);
    });
  });

  describe("labels", () => {
    it("removing INBOX archives the thread", async () => {
      await provider.removeLabel("conv-1", "INBOX");
      expect(moveMessages).toHaveBeenCalledWith(client, ["m-inbox"], "archive");
    });

    it("adding INBOX brings archived mail back", async () => {
      await provider.addLabel("conv-1", "INBOX");
      expect(moveMessages).toHaveBeenCalledWith(client, ["m-archived"], "inbox");
    });

    it("adding a folder label moves the thread there", async () => {
      await provider.addLabel("conv-1", "graph-folder:f-jobs");
      expect(moveMessages).toHaveBeenCalledWith(client, ["m-inbox", "m-archived"], "f-jobs");
    });

    it("adding a category label appends the Outlook category where missing", async () => {
      client.batch
        .mockResolvedValueOnce([
          { id: "0", status: 200, body: { categories: [] } },
          { id: "1", status: 200, body: { categories: ["Blue category"] } },
          { id: "2", status: 200, body: { categories: ["Red category"] } },
        ])
        .mockResolvedValueOnce([{ id: "0", status: 200 }, { id: "1", status: 200 }]);

      await provider.addLabel("conv-1", "graph-cat:Blue category");

      const updates = client.batch.mock.calls[1]![0];
      expect(updates.map((r) => [r.url, r.body])).toEqual([
        ["/me/messages/m-inbox", { categories: ["Blue category"] }],
        ["/me/messages/m-archived", { categories: ["Red category", "Blue category"] }],
      ]);
    });

    it("removing a category label strips only that category", async () => {
      mockDb.select.mockResolvedValue([{ id: "m1", imap_folder: "f-inbox", date: 1 }]);
      client.batch
        .mockResolvedValueOnce([{ id: "0", status: 200, body: { categories: ["Blue category", "Red category"] } }])
        .mockResolvedValueOnce([{ id: "0", status: 200 }]);

      await provider.removeLabel("conv-1", "graph-cat:Blue category");

      expect(client.batch.mock.calls[1]![0][0]!.body).toEqual({ categories: ["Red category"] });
    });

    it("removing a folder label is a no-op", async () => {
      const warn = vi.spyOn(console, "warn").mockImplementation(() => undefined);
      await provider.removeLabel("conv-1", "graph-folder:f-jobs");
      expect(moveMessages).not.toHaveBeenCalled();
      expect(client.batch).not.toHaveBeenCalled();
      warn.mockRestore();
    });
  });

  describe("send and drafts", () => {
    it("sends MIME as standard base64 text", async () => {
      vi.useFakeTimers();
      const result = await provider.sendMessage("a-b_cw");
      expect(client.request).toHaveBeenCalledWith("/me/sendMail", {
        method: "POST",
        contentType: "text/plain",
        body: "a+b/cw==",
      });
      expect(result.id).toMatch(/^graph-sent-/);
      vi.useRealTimers();
    });

    it("creates a draft from MIME and returns its ID", async () => {
      client.request.mockResolvedValue({ id: "draft-1" });
      await expect(provider.createDraft("TUlNRQ")).resolves.toEqual({ draftId: "draft-1" });
      expect(client.request).toHaveBeenCalledWith("/me/messages", {
        method: "POST",
        contentType: "text/plain",
        body: "TUlNRQ==",
      });
    });

    it("updateDraft creates the replacement before deleting the old draft", async () => {
      client.request.mockResolvedValueOnce({ id: "draft-2" }).mockResolvedValueOnce(undefined);
      await expect(provider.updateDraft("draft-1", "TUlNRQ")).resolves.toEqual({ draftId: "draft-2" });
      expect(client.request.mock.calls[1]).toEqual(["/me/messages/draft-1", { method: "DELETE" }]);
    });

    it("updateDraft keeps the new draft even if deleting the old one fails", async () => {
      const warn = vi.spyOn(console, "warn").mockImplementation(() => undefined);
      client.request.mockResolvedValueOnce({ id: "draft-2" }).mockRejectedValueOnce(new Error("gone"));
      await expect(provider.updateDraft("draft-1", "TUlNRQ")).resolves.toEqual({ draftId: "draft-2" });
      warn.mockRestore();
    });
  });

  describe("message reads", () => {
    it("fetchAttachment returns base64 content and size", async () => {
      client.requestBytes.mockResolvedValue(new Uint8Array([72, 105]));
      await expect(provider.fetchAttachment("m1", "att-1")).resolves.toEqual({ data: "SGk=", size: 2 });
      expect(client.requestBytes).toHaveBeenCalledWith("/me/messages/m1/attachments/att-1/$value");
    });

    it("fetchRawMessage returns the MIME source", async () => {
      client.requestText.mockResolvedValue("From: a@b.c\r\n\r\nbody");
      await expect(provider.fetchRawMessage("m1")).resolves.toBe("From: a@b.c\r\n\r\nbody");
      expect(client.requestText).toHaveBeenCalledWith("/me/messages/m1/$value");
    });

    it("fetchMessage labels the message with its folder", async () => {
      client.request.mockResolvedValue({ id: "m1", conversationId: "conv-1", parentFolderId: "f-archive", isRead: true });
      const parsed = await provider.fetchMessage("m1");
      expect(parsed.labelIds).toEqual(["archive"]);
    });
  });

  describe("folders", () => {
    it("listFolders maps the folder context", async () => {
      vi.mocked(loadFolderContext).mockResolvedValue({
        folders: [
          { folderId: "f-inbox", labelId: "INBOX", labelName: "Inbox", type: "system", specialUse: null, totalItemCount: 5, unreadItemCount: 1 },
          { folderId: "f-937", labelId: "graph-folder:f-937", labelName: "Jobs/937 View St", type: "user", specialUse: null, totalItemCount: 2, unreadItemCount: 0 },
        ],
        labelByFolderId: new Map(),
        wellKnownIds: {},
      });

      const folders = await provider.listFolders();

      expect(folders[1]).toEqual({
        id: "graph-folder:f-937",
        name: "937 View St",
        path: "Jobs/937 View St",
        type: "user",
        specialUse: null,
        delimiter: "/",
        messageCount: 2,
        unreadCount: 0,
      });
    });

    it("createFolder creates a child folder under a parent label", async () => {
      client.request.mockResolvedValue({ id: "f-new", displayName: "Closeout" });
      const folder = await provider.createFolder("Closeout", "graph-folder:f-jobs");
      expect(client.request).toHaveBeenCalledWith("/me/mailFolders/f-jobs/childFolders", {
        method: "POST",
        body: { displayName: "Closeout" },
      });
      expect(folder.id).toBe("graph-folder:f-new");
    });

    it("renames and deletes folders by label ID", async () => {
      await provider.renameFolder("graph-folder:f-jobs", "Projects");
      await provider.deleteFolder("graph-folder:f-jobs");
      expect(client.request).toHaveBeenNthCalledWith(1, "/me/mailFolders/f-jobs", { method: "PATCH", body: { displayName: "Projects" } });
      expect(client.request).toHaveBeenNthCalledWith(2, "/me/mailFolders/f-jobs", { method: "DELETE" });
    });
  });

  describe("sync and connection", () => {
    it("delegates sync to the Graph sync module", async () => {
      await provider.initialSync(90);
      await provider.deltaSync("ignored");
      expect(graphInitialSync).toHaveBeenCalledWith("acc-1", 90, expect.any(Function));
      expect(graphDeltaSync).toHaveBeenCalledWith("acc-1");
    });

    it("testConnection reports the signed-in mailbox", async () => {
      client.request.mockResolvedValue({ mail: "jeremy@example.com", displayName: "Jeremy" });
      await expect(provider.testConnection()).resolves.toEqual({ success: true, message: "Connected as jeremy@example.com" });
    });

    it("testConnection reports failures", async () => {
      client.request.mockRejectedValue(new Error("Forbidden"));
      await expect(provider.testConnection()).resolves.toEqual({ success: false, message: "Forbidden" });
    });
  });
});
