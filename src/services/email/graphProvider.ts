import type { EmailProvider, EmailFolder, SyncResult } from "./types";
import type { ParsedMessage } from "../gmail/messageParser";
import { getDb } from "../db/connection";
import { getLabelsForAccount } from "../db/labels";
import { GraphClient, assertBatchOk, type GraphBatchRequest } from "../graph/graphClient";
import {
  GRAPH_FOLDER_LABEL_PREFIX,
  base64UrlToBase64,
  bytesToBase64,
  categoryNameFromLabel,
  graphMessageToParsed,
  isCategoryLabel,
  type GraphMailFolder,
  type GraphMessage,
} from "../graph/graphMapper";
import { graphDeltaSync, graphInitialSync, loadFolderContext, moveMessages } from "../graph/graphSync";

/** System label → well-known Exchange folder name (accepted wherever Graph takes a folder ID). */
const LABEL_TO_WELL_KNOWN: Record<string, string> = {
  INBOX: "inbox",
  SENT: "sentitems",
  DRAFT: "drafts",
  TRASH: "deleteditems",
  SPAM: "junkemail",
  archive: "archive",
};

const HTML_BODY_PREFERENCE = 'outlook.body-content-type="html"';

const FETCH_MESSAGE_QUERY =
  "$select=id,conversationId,parentFolderId,subject,bodyPreview,body,from,sender,toRecipients,ccRecipients,bccRecipients,replyTo,receivedDateTime,sentDateTime,createdDateTime,isRead,isDraft,importance,flag,hasAttachments,internetMessageId,internetMessageHeaders,categories" +
  "&$expand=attachments($select=id,name,contentType,size,isInline)";

interface ThreadMessageRow {
  id: string;
  /** Label ID of the folder the message is in, when known. */
  folderLabel: string | null;
  date: number;
}

/**
 * EmailProvider for Microsoft 365 / Exchange Online mailboxes via Microsoft Graph.
 *
 * Exchange keeps each message in exactly one folder, so Gmail-style thread
 * actions are translated into per-message moves and updates. Outlook
 * categories are surfaced as labels (`graph-cat:{name}`).
 */
export class GraphProvider implements EmailProvider {
  readonly accountId: string;
  readonly type = "graph" as const;
  private client: GraphClient;

  constructor(accountId: string, client: GraphClient) {
    this.accountId = accountId;
    this.client = client;
  }

  // ---- Folder/Label operations ----

  async listFolders(): Promise<EmailFolder[]> {
    const context = await loadFolderContext(this.client);
    return context.folders.map((folder) => ({
      id: folder.labelId,
      name: folder.labelName.split("/").pop() ?? folder.labelName,
      path: folder.labelName,
      type: folder.type,
      specialUse: folder.specialUse,
      delimiter: "/",
      messageCount: folder.totalItemCount,
      unreadCount: folder.unreadItemCount,
    }));
  }

  async createFolder(name: string, parentPath?: string): Promise<EmailFolder> {
    const url = parentPath
      ? `/me/mailFolders/${encodeURIComponent(this.resolveFolderId(parentPath))}/childFolders`
      : "/me/mailFolders";
    const folder = await this.client.request<GraphMailFolder>(url, {
      method: "POST",
      body: { displayName: name },
    });
    return {
      id: `${GRAPH_FOLDER_LABEL_PREFIX}${folder.id}`,
      name: folder.displayName,
      path: folder.displayName,
      type: "user",
      specialUse: null,
      delimiter: "/",
      messageCount: 0,
      unreadCount: 0,
    };
  }

  async deleteFolder(path: string): Promise<void> {
    await this.client.request(`/me/mailFolders/${encodeURIComponent(this.resolveFolderId(path))}`, {
      method: "DELETE",
    });
  }

  async renameFolder(path: string, newName: string): Promise<void> {
    await this.client.request(`/me/mailFolders/${encodeURIComponent(this.resolveFolderId(path))}`, {
      method: "PATCH",
      body: { displayName: newName },
    });
  }

  // ---- Sync operations ----
  // The sync modules write straight to the local DB (same pattern as IMAP),
  // so these return empty message lists.

  async initialSync(
    daysBack: number,
    onProgress?: (phase: string, current: number, total: number) => void,
  ): Promise<SyncResult> {
    await graphInitialSync(this.accountId, daysBack, (p) => onProgress?.(p.phase, p.current, p.total));
    return { messages: [] };
  }

  async deltaSync(_syncToken: string): Promise<SyncResult> {
    await graphDeltaSync(this.accountId);
    return { messages: [] };
  }

  // ---- Message operations ----

  async fetchMessage(messageId: string): Promise<ParsedMessage> {
    const msg = await this.client.request<GraphMessage>(
      `/me/messages/${encodeURIComponent(messageId)}?${FETCH_MESSAGE_QUERY}`,
      { prefer: [HTML_BODY_PREFERENCE] },
    );
    const folderLabels = await this.getFolderLabelMap();
    const folderLabel = (msg.parentFolderId && folderLabels.get(msg.parentFolderId)) || "INBOX";
    return graphMessageToParsed(msg, folderLabel);
  }

  async fetchAttachment(
    messageId: string,
    attachmentId: string,
  ): Promise<{ data: string; size: number }> {
    const bytes = await this.client.requestBytes(
      `/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}/$value`,
    );
    return { data: bytesToBase64(bytes), size: bytes.length };
  }

  async fetchRawMessage(messageId: string): Promise<string> {
    return this.client.requestText(`/me/messages/${encodeURIComponent(messageId)}/$value`);
  }

  // ---- Actions ----

  async archive(threadId: string, messageIds: string[]): Promise<void> {
    const rows = await this.getThreadMessages(threadId, messageIds);
    await this.move(this.idsIn(rows, ["INBOX"]), "archive");
  }

  async trash(threadId: string, messageIds: string[]): Promise<void> {
    const rows = await this.getThreadMessages(threadId, messageIds);
    await this.move(this.idsNotIn(rows, ["TRASH"]), "TRASH");
  }

  async permanentDelete(threadId: string, messageIds: string[]): Promise<void> {
    const rows = await this.getThreadMessages(threadId, messageIds);
    await this.runBatch(
      rows.map((row) => ({ method: "DELETE" as const, url: `/me/messages/${encodeURIComponent(row.id)}` })),
    );
  }

  async markRead(threadId: string, messageIds: string[], read: boolean): Promise<void> {
    const rows = await this.getThreadMessages(threadId, messageIds);
    await this.patchMessages(rows.map((r) => r.id), { isRead: read });
  }

  async star(threadId: string, messageIds: string[], starred: boolean): Promise<void> {
    const rows = await this.getThreadMessages(threadId, messageIds);
    if (rows.length === 0) return;

    if (starred) {
      // Each flagged message becomes a To Do item in Outlook, so flag only the latest one
      const latest = rows.reduce((a, b) => (b.date >= a.date ? b : a));
      await this.patchMessages([latest.id], { flag: { flagStatus: "flagged" } });
    } else {
      await this.patchMessages(rows.map((r) => r.id), { flag: { flagStatus: "notFlagged" } });
    }
  }

  async spam(threadId: string, messageIds: string[], isSpam: boolean): Promise<void> {
    const rows = await this.getThreadMessages(threadId, messageIds);
    if (isSpam) {
      await this.move(this.idsNotIn(rows, ["SPAM", "SENT", "DRAFT"]), "SPAM");
    } else {
      await this.move(this.idsIn(rows, ["SPAM"]), "INBOX");
    }
  }

  async moveToFolder(threadId: string, messageIds: string[], folderPath: string): Promise<void> {
    const rows = await this.getThreadMessages(threadId, messageIds);
    // Sent copies and drafts stay where they are, like moving a conversation in Outlook
    await this.move(this.idsNotIn(rows, [folderPath, "SENT", "DRAFT"]), folderPath);
  }

  async addLabel(threadId: string, labelId: string): Promise<void> {
    if (isCategoryLabel(labelId)) {
      await this.updateCategories(threadId, categoryNameFromLabel(labelId), true);
      return;
    }

    switch (labelId) {
      case "INBOX": {
        // "Add INBOX" is how Gmail un-archives a thread
        const rows = await this.getThreadMessages(threadId, []);
        await this.move(this.idsIn(rows, ["archive", "TRASH", "SPAM"]), "INBOX");
        return;
      }
      case "TRASH":
        return this.trash(threadId, []);
      case "SPAM":
        return this.spam(threadId, [], true);
      case "STARRED":
        return this.star(threadId, [], true);
      case "UNREAD":
        return this.markRead(threadId, [], false);
      case "IMPORTANT": {
        const rows = await this.getThreadMessages(threadId, []);
        await this.patchMessages(rows.map((r) => r.id), { importance: "high" });
        return;
      }
      default:
        // Any other label is a folder
        return this.moveToFolder(threadId, [], labelId);
    }
  }

  async removeLabel(threadId: string, labelId: string): Promise<void> {
    if (isCategoryLabel(labelId)) {
      await this.updateCategories(threadId, categoryNameFromLabel(labelId), false);
      return;
    }

    switch (labelId) {
      case "INBOX":
        // "Remove INBOX" is how Gmail archives a thread
        return this.archive(threadId, []);
      case "STARRED":
        return this.star(threadId, [], false);
      case "UNREAD":
        return this.markRead(threadId, [], true);
      case "SPAM":
        return this.spam(threadId, [], false);
      case "IMPORTANT": {
        const rows = await this.getThreadMessages(threadId, []);
        await this.patchMessages(rows.map((r) => r.id), { importance: "normal" });
        return;
      }
      default:
        // A message always lives in exactly one folder, so there is nothing to
        // "remove" — moving it elsewhere is what takes it out of this folder.
        console.warn(`[graph] removeLabel(${labelId}) ignored: Exchange folders can't be removed from a message`);
    }
  }

  // ---- Send/Draft operations ----

  async sendMessage(rawBase64Url: string, _threadId?: string): Promise<{ id: string }> {
    // Graph accepts a complete MIME message (base64) and threads replies from its headers
    await this.client.request("/me/sendMail", {
      method: "POST",
      contentType: "text/plain",
      body: base64UrlToBase64(rawBase64Url),
    });

    // sendMail returns no ID; pull the Sent Items copy in shortly instead of waiting for the next cycle
    setTimeout(() => {
      import("../gmail/syncManager")
        .then(({ syncAccount }) => syncAccount(this.accountId))
        .catch((err) => console.warn("[graph] Post-send sync failed:", err));
    }, 3000);

    return { id: `graph-sent-${Date.now()}` };
  }

  async createDraft(rawBase64Url: string, _threadId?: string): Promise<{ draftId: string }> {
    const draft = await this.client.request<GraphMessage>("/me/messages", {
      method: "POST",
      contentType: "text/plain",
      body: base64UrlToBase64(rawBase64Url),
    });
    return { draftId: draft.id };
  }

  async updateDraft(
    draftId: string,
    rawBase64Url: string,
    threadId?: string,
  ): Promise<{ draftId: string }> {
    // A draft can't be replaced with new MIME content in place, so swap it for a fresh one
    const created = await this.createDraft(rawBase64Url, threadId);
    try {
      await this.deleteDraft(draftId);
    } catch (err) {
      console.warn(`[graph] Failed to delete superseded draft ${draftId}:`, err);
    }
    return created;
  }

  async deleteDraft(draftId: string): Promise<void> {
    await this.client.request(`/me/messages/${encodeURIComponent(draftId)}`, { method: "DELETE" });
  }

  // ---- Connection ----

  async testConnection(): Promise<{ success: boolean; message: string }> {
    try {
      const profile = await this.getProfile();
      return { success: true, message: `Connected as ${profile.email}` };
    } catch (err) {
      return {
        success: false,
        message: err instanceof Error ? err.message : "Unknown connection error",
      };
    }
  }

  async getProfile(): Promise<{ email: string; name?: string }> {
    const me = await this.client.request<{
      mail?: string | null;
      userPrincipalName?: string | null;
      displayName?: string | null;
    }>("/me?$select=mail,userPrincipalName,displayName");
    return {
      email: me.mail ?? me.userPrincipalName ?? "",
      name: me.displayName ?? undefined,
    };
  }

  // ---- Helpers ----

  /** Graph folder ID → label ID, from the locally synced labels. */
  private async getFolderLabelMap(): Promise<Map<string, string>> {
    const labels = await getLabelsForAccount(this.accountId);
    const map = new Map<string, string>();
    for (const label of labels) {
      if (label.imap_folder_path) map.set(label.imap_folder_path, label.id);
    }
    return map;
  }

  /**
   * Messages an action applies to: the given IDs, or the whole thread when
   * none are given. Each row carries the label of the folder it lives in.
   */
  private async getThreadMessages(threadId: string, messageIds: string[]): Promise<ThreadMessageRow[]> {
    const db = await getDb();
    const stored = await db.select<{ id: string; imap_folder: string | null; date: number }[]>(
      "SELECT id, imap_folder, date FROM messages WHERE account_id = $1 AND thread_id = $2",
      [this.accountId, threadId],
    );
    const folderLabels = await this.getFolderLabelMap();
    const rows: ThreadMessageRow[] = stored.map((row) => ({
      id: row.id,
      folderLabel: row.imap_folder ? (folderLabels.get(row.imap_folder) ?? null) : null,
      date: row.date,
    }));

    if (messageIds.length === 0) return rows;

    const byId = new Map(rows.map((r) => [r.id, r]));
    return messageIds.map((id) => byId.get(id) ?? { id, folderLabel: null, date: 0 });
  }

  private idsIn(rows: ThreadMessageRow[], labels: string[]): string[] {
    return rows.filter((r) => r.folderLabel !== null && labels.includes(r.folderLabel)).map((r) => r.id);
  }

  private idsNotIn(rows: ThreadMessageRow[], labels: string[]): string[] {
    return rows.filter((r) => r.folderLabel === null || !labels.includes(r.folderLabel)).map((r) => r.id);
  }

  /** Turn a label ID (or raw folder ID) into something Graph accepts as a folder ID. */
  private resolveFolderId(labelOrFolderId: string): string {
    if (labelOrFolderId.startsWith(GRAPH_FOLDER_LABEL_PREFIX)) {
      return labelOrFolderId.slice(GRAPH_FOLDER_LABEL_PREFIX.length);
    }
    return LABEL_TO_WELL_KNOWN[labelOrFolderId] ?? labelOrFolderId;
  }

  /**
   * Move messages to the folder behind `destinationLabel` and record the new
   * folder locally, so follow-up actions don't have to wait for the next sync.
   */
  private async move(messageIds: string[], destinationLabel: string): Promise<void> {
    if (messageIds.length === 0) return;
    await moveMessages(this.client, messageIds, this.resolveFolderId(destinationLabel));

    try {
      const folderLabels = await this.getFolderLabelMap();
      const folderId = [...folderLabels].find(([, labelId]) => labelId === destinationLabel)?.[0];
      if (!folderId) return;
      const db = await getDb();
      const placeholders = messageIds.map((_, i) => `$${i + 3}`).join(", ");
      await db.execute(
        `UPDATE messages SET imap_folder = $1 WHERE account_id = $2 AND id IN (${placeholders})`,
        [folderId, this.accountId, ...messageIds],
      );
    } catch (err) {
      // The server move succeeded; the next sync corrects the local folder
      console.warn("[graph] Failed to record new folder locally:", err);
    }
  }

  private async runBatch(requests: Omit<GraphBatchRequest, "id">[]): Promise<void> {
    if (requests.length === 0) return;
    const responses = await this.client.batch(requests.map((req, i) => ({ ...req, id: String(i) })));
    assertBatchOk(responses);
  }

  private async patchMessages(messageIds: string[], body: Record<string, unknown>): Promise<void> {
    await this.runBatch(
      messageIds.map((id) => ({
        method: "PATCH" as const,
        url: `/me/messages/${encodeURIComponent(id)}`,
        body,
      })),
    );
  }

  /** Add or remove an Outlook category on every message of a thread. */
  private async updateCategories(threadId: string, category: string, add: boolean): Promise<void> {
    const rows = await this.getThreadMessages(threadId, []);
    if (rows.length === 0) return;

    // Categories are a full-replace list, so read the current values first
    const current = await this.client.batch(
      rows.map((row, i) => ({
        id: String(i),
        method: "GET" as const,
        url: `/me/messages/${encodeURIComponent(row.id)}?$select=categories`,
      })),
    );

    const updates: Omit<GraphBatchRequest, "id">[] = [];
    for (const res of current) {
      if (res.status !== 200) continue;
      const row = rows[Number(res.id)];
      if (!row) continue;
      const existing = (res.body as { categories?: string[] } | undefined)?.categories ?? [];
      const has = existing.includes(category);
      if (add === has) continue;
      updates.push({
        method: "PATCH",
        url: `/me/messages/${encodeURIComponent(row.id)}`,
        body: { categories: add ? [...existing, category] : existing.filter((c) => c !== category) },
      });
    }
    await this.runBatch(updates);
  }
}
