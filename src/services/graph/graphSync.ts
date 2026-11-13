import { getDb, withTransaction } from "../db/connection";
import { updateAccountSyncState } from "../db/accounts";
import { upsertLabel, getLabelsForAccount, deleteLabel } from "../db/labels";
import { upsertThread, setThreadLabels, deleteThread, getMutedThreadIds } from "../db/threads";
import { upsertMessage, deleteMessage } from "../db/messages";
import { upsertAttachment } from "../db/attachments";
import { getSetting, setSetting } from "../db/settings";
import { getPendingOpsForResource } from "../db/pendingOperations";
import { getThreadCategory } from "../db/threadCategories";
import { getVipSenders } from "../db/notificationVips";
import { shouldNotifyForMessage, queueNewEmailNotification } from "../notifications/notificationManager";
import { applyFiltersToMessages } from "../filters/filterEngine";
import { GraphClient, GraphError, getGraphClient, assertBatchOk } from "./graphClient";
import {
  WELL_KNOWN_LABELS,
  SKIPPED_WELL_KNOWN_FOLDERS,
  GRAPH_FOLDER_LABEL_PREFIX,
  GRAPH_CATEGORY_LABEL_PREFIX,
  categoryColor,
  categoryLabelId,
  graphMessageToParsed,
  mapFoldersToLabels,
  odataQuote,
  type GraphCategory,
  type GraphFolderLabel,
  type GraphMailFolder,
  type GraphMessage,
  type GraphParsedMessage,
} from "./graphMapper";

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const MESSAGE_PAGE_SIZE = 50;
const DELTA_PAGE_SIZE = 500;
const FOLDER_CONCURRENCY = 4;
const THREAD_CONCURRENCY = 4;

const HTML_BODY_PREFERENCE = 'outlook.body-content-type="html"';

const MESSAGE_SELECT = [
  "id",
  "conversationId",
  "parentFolderId",
  "subject",
  "bodyPreview",
  "body",
  "from",
  "sender",
  "toRecipients",
  "ccRecipients",
  "bccRecipients",
  "replyTo",
  "receivedDateTime",
  "sentDateTime",
  "createdDateTime",
  "isRead",
  "isDraft",
  "importance",
  "flag",
  "hasAttachments",
  "internetMessageId",
  "internetMessageHeaders",
  "categories",
].join(",");

/** Delta results only need enough to know which conversation changed. */
const DELTA_SELECT = "conversationId,isRead,parentFolderId";

const ATTACHMENT_FIELDS = "id,name,contentType,size,isInline";

/**
 * `contentId` lives on the fileAttachment subtype, so selecting it needs a
 * type cast. If the service rejects the cast we fall back to the base fields
 * (inline images then show as regular attachments instead of in the body).
 */
let selectAttachmentContentId = true;

function messageQuery(): string {
  const fields = selectAttachmentContentId
    ? `${ATTACHMENT_FIELDS},microsoft.graph.fileAttachment/contentId`
    : ATTACHMENT_FIELDS;
  return `$select=${MESSAGE_SELECT}&$expand=attachments($select=${fields})`;
}

async function withAttachmentFallback<T>(run: () => Promise<T>): Promise<T> {
  try {
    return await run();
  } catch (err) {
    if (selectAttachmentContentId && err instanceof GraphError && err.status === 400) {
      console.warn("[graphSync] Attachment contentId select rejected — retrying without it:", err.message);
      selectAttachmentContentId = false;
      return run();
    }
    throw err;
  }
}

/** Test hook: restore the default attachment select. */
export function resetGraphSyncFlags(): void {
  selectAttachmentContentId = true;
}

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface GraphSyncProgress {
  phase: "labels" | "threads" | "messages" | "done";
  current: number;
  total: number;
}

export type GraphSyncProgressCallback = (progress: GraphSyncProgress) => void;

export interface GraphFolderContext {
  folders: GraphFolderLabel[];
  /** Graph folder ID → Velo label ID. */
  labelByFolderId: Map<string, string>;
  /** Well-known name (inbox, archive, …) → Graph folder ID. */
  wellKnownIds: Record<string, string>;
}

interface GraphSyncState {
  /** Graph folder ID → delta link for that folder's messages. */
  deltaLinks: Record<string, string>;
  /** Conversations whose refresh failed and must be retried next cycle. */
  retryThreadIds: string[];
}

/** Running per-thread summary, built up message by message. */
interface ThreadAggregate {
  threadId: string;
  subject: string | null;
  firstDate: number;
  snippet: string;
  lastDate: number;
  lastFromAddress: string | null;
  lastListUnsubscribe: string | null;
  messageIds: Set<string>;
  inboxMessageIds: Set<string>;
  hasUnread: boolean;
  isStarred: boolean;
  hasAttachments: boolean;
  labelIds: Set<string>;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

async function parallelLimit<T>(items: T[], limit: number, worker: (item: T) => Promise<void>): Promise<void> {
  let index = 0;
  const run = async (): Promise<void> => {
    while (index < items.length) {
      const item = items[index++]!;
      await worker(item);
    }
  };
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, run));
}

function syncStateKey(accountId: string): string {
  return `graph_sync_state:${accountId}`;
}

async function loadSyncState(accountId: string): Promise<GraphSyncState> {
  const raw = await getSetting(syncStateKey(accountId));
  if (!raw) return { deltaLinks: {}, retryThreadIds: [] };
  try {
    const parsed = JSON.parse(raw) as Partial<GraphSyncState>;
    return { deltaLinks: parsed.deltaLinks ?? {}, retryThreadIds: parsed.retryThreadIds ?? [] };
  } catch {
    return { deltaLinks: {}, retryThreadIds: [] };
  }
}

async function saveSyncState(accountId: string, state: GraphSyncState): Promise<void> {
  await setSetting(syncStateKey(accountId), JSON.stringify(state));
}

function isDeltaExpired(err: unknown): boolean {
  if (!(err instanceof GraphError)) return false;
  if (err.status === 410) return true;
  const code = err.code?.toLowerCase() ?? "";
  return code.includes("syncstate") || code.includes("resyncrequired");
}

function sinceIso(daysBack: number): string {
  return new Date(Date.now() - daysBack * 24 * 60 * 60 * 1000).toISOString();
}

// ---------------------------------------------------------------------------
// Folders and categories → labels
// ---------------------------------------------------------------------------

async function listAllFolders(client: GraphClient): Promise<GraphMailFolder[]> {
  type FolderWithChildren = GraphMailFolder & { childFolderCount?: number };
  const select = "$select=id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount&$top=100";

  const all: FolderWithChildren[] = [];
  let pending = await client.getAll<FolderWithChildren>(`/me/mailFolders?${select}`);

  // Breadth-first walk of the hierarchy
  for (let depth = 0; pending.length > 0 && depth < 20; depth++) {
    all.push(...pending);
    const parents = pending.filter((f) => (f.childFolderCount ?? 0) > 0);
    const next: FolderWithChildren[] = [];
    await parallelLimit(parents, FOLDER_CONCURRENCY, async (parent) => {
      const children = await client.getAll<FolderWithChildren>(
        `/me/mailFolders/${encodeURIComponent(parent.id)}/childFolders?${select}`,
      );
      next.push(...children);
    });
    pending = next;
  }

  return all;
}

async function resolveWellKnownIds(client: GraphClient): Promise<Record<string, string>> {
  const names = [...Object.keys(WELL_KNOWN_LABELS), ...SKIPPED_WELL_KNOWN_FOLDERS];
  const responses = await client.batch(
    names.map((name) => ({ id: name, method: "GET" as const, url: `/me/mailFolders/${name}?$select=id` })),
  );

  const ids: Record<string, string> = {};
  for (const res of responses) {
    // Not every mailbox has every well-known folder — missing ones are simply skipped
    const id = (res.body as { id?: string } | undefined)?.id;
    if (res.status === 200 && id) ids[res.id] = id;
  }
  if (!ids.inbox) {
    throw new Error("Could not find the Inbox folder for this Microsoft 365 mailbox.");
  }
  return ids;
}

/** Read the folder hierarchy from Graph without touching the DB. */
export async function loadFolderContext(client: GraphClient): Promise<GraphFolderContext> {
  const [folders, wellKnownIds] = await Promise.all([listAllFolders(client), resolveWellKnownIds(client)]);
  const mapped = mapFoldersToLabels(folders, wellKnownIds);
  return {
    folders: mapped,
    labelByFolderId: new Map(mapped.map((f) => [f.folderId, f.labelId])),
    wellKnownIds,
  };
}

/**
 * Mirror mail folders and Outlook categories into the labels table.
 * Folders become folder labels; categories become user labels.
 */
export async function syncFoldersAndCategories(
  client: GraphClient,
  accountId: string,
): Promise<GraphFolderContext> {
  const context = await loadFolderContext(client);

  const currentLabelIds = new Set<string>();
  for (const folder of context.folders) {
    currentLabelIds.add(folder.labelId);
    await upsertLabel({
      id: folder.labelId,
      accountId,
      name: folder.labelName,
      type: folder.type,
      imapFolderPath: folder.folderId,
      imapSpecialUse: folder.specialUse,
    });
  }

  let categoriesLoaded = false;
  try {
    const categories = await client.getAll<GraphCategory>("/me/outlook/masterCategories");
    for (const category of categories) {
      const labelId = categoryLabelId(category.displayName);
      currentLabelIds.add(labelId);
      await upsertLabel({
        id: labelId,
        accountId,
        name: category.displayName,
        type: "user",
        colorBg: categoryColor(category.color),
        colorFg: categoryColor(category.color) ? "#ffffff" : null,
      });
    }
    categoriesLoaded = true;
  } catch (err) {
    // Categories are optional (needs MailboxSettings.Read) — mail still syncs without them
    console.warn("[graphSync] Could not load Outlook categories:", err);
  }

  // Remove labels for folders/categories that no longer exist on the server
  const existing = await getLabelsForAccount(accountId);
  for (const label of existing) {
    const isFolderLabel = label.id.startsWith(GRAPH_FOLDER_LABEL_PREFIX);
    const isCategory = label.id.startsWith(GRAPH_CATEGORY_LABEL_PREFIX);
    if (currentLabelIds.has(label.id)) continue;
    if (isFolderLabel || (isCategory && categoriesLoaded)) {
      await deleteLabel(accountId, label.id);
    }
  }

  return context;
}

// ---------------------------------------------------------------------------
// Storage
// ---------------------------------------------------------------------------

function mergeIntoAggregate(
  aggregates: Map<string, ThreadAggregate>,
  msg: GraphParsedMessage,
): ThreadAggregate {
  let agg = aggregates.get(msg.threadId);
  if (!agg) {
    agg = {
      threadId: msg.threadId,
      subject: msg.subject,
      firstDate: msg.date,
      snippet: msg.snippet,
      lastDate: msg.date,
      lastFromAddress: msg.fromAddress,
      lastListUnsubscribe: msg.listUnsubscribe,
      messageIds: new Set(),
      inboxMessageIds: new Set(),
      hasUnread: false,
      isStarred: false,
      hasAttachments: false,
      labelIds: new Set(),
    };
    aggregates.set(msg.threadId, agg);
  }

  if (msg.date <= agg.firstDate) {
    agg.firstDate = msg.date;
    agg.subject = msg.subject;
  }
  if (msg.date >= agg.lastDate) {
    agg.lastDate = msg.date;
    agg.snippet = msg.snippet;
    agg.lastFromAddress = msg.fromAddress;
    agg.lastListUnsubscribe = msg.listUnsubscribe;
  }

  agg.messageIds.add(msg.id);
  if (msg.labelIds.includes("INBOX")) agg.inboxMessageIds.add(msg.id);
  if (!msg.isRead) agg.hasUnread = true;
  if (msg.isStarred) agg.isStarred = true;
  if (msg.hasAttachments) agg.hasAttachments = true;
  for (const labelId of msg.labelIds) agg.labelIds.add(labelId);

  return agg;
}

async function writeThread(accountId: string, agg: ThreadAggregate): Promise<void> {
  await upsertThread({
    id: agg.threadId,
    accountId,
    subject: agg.subject,
    snippet: agg.snippet,
    lastMessageAt: agg.lastDate,
    messageCount: agg.messageIds.size,
    isRead: !agg.hasUnread,
    isStarred: agg.isStarred,
    isImportant: agg.labelIds.has("IMPORTANT"),
    hasAttachments: agg.hasAttachments,
  });
  await setThreadLabels(accountId, agg.threadId, [...agg.labelIds]);
}

async function writeMessage(accountId: string, msg: GraphParsedMessage): Promise<void> {
  await upsertMessage({
    id: msg.id,
    accountId,
    threadId: msg.threadId,
    fromAddress: msg.fromAddress,
    fromName: msg.fromName,
    toAddresses: msg.toAddresses,
    ccAddresses: msg.ccAddresses,
    bccAddresses: msg.bccAddresses,
    replyTo: msg.replyTo,
    subject: msg.subject,
    snippet: msg.snippet,
    date: msg.date,
    isRead: msg.isRead,
    isStarred: msg.isStarred,
    bodyHtml: msg.bodyHtml,
    bodyText: msg.bodyText,
    rawSize: msg.rawSize,
    internalDate: msg.internalDate,
    listUnsubscribe: msg.listUnsubscribe,
    listUnsubscribePost: msg.listUnsubscribePost,
    authResults: msg.authResults,
    messageIdHeader: msg.messageIdHeader,
    referencesHeader: msg.referencesHeader,
    inReplyToHeader: msg.inReplyToHeader,
    // Graph accounts reuse this column for the Graph folder ID the message lives in
    imapFolder: msg.folderId,
  });

  for (const att of msg.attachments) {
    await upsertAttachment({
      id: `${msg.id}_${att.gmailAttachmentId}`,
      messageId: msg.id,
      accountId,
      filename: att.filename,
      mimeType: att.mimeType,
      size: att.size,
      gmailAttachmentId: att.gmailAttachmentId,
      contentId: att.contentId,
      isInline: att.isInline,
    });
  }
}

/**
 * Rule-based categorisation for inbox threads, mirroring the Gmail sync:
 * respects manual categories, optional auto-archive and bundle delivery holds.
 */
async function categorizeThread(
  client: GraphClient,
  accountId: string,
  agg: ThreadAggregate,
  context: GraphFolderContext,
  autoArchiveCategories: Set<string>,
): Promise<void> {
  if (!agg.labelIds.has("INBOX")) return;

  const { getThreadCategoryWithManual, setThreadCategory } = await import("@/services/db/threadCategories");
  const existing = await getThreadCategoryWithManual(accountId, agg.threadId);
  if (existing?.isManual) return;

  const { categorizeByRules } = await import("@/services/categorization/ruleEngine");
  const category = categorizeByRules({
    labelIds: [...agg.labelIds],
    fromAddress: agg.lastFromAddress,
    listUnsubscribe: agg.lastListUnsubscribe,
  });
  await setThreadCategory(accountId, agg.threadId, category, false);

  if (category === "Primary") return;

  if (autoArchiveCategories.has(category) && context.wellKnownIds.archive) {
    try {
      await moveMessages(client, [...agg.inboxMessageIds], context.wellKnownIds.archive);
      agg.labelIds.delete("INBOX");
      agg.labelIds.add("archive");
      agg.inboxMessageIds.clear();
      await setThreadLabels(accountId, agg.threadId, [...agg.labelIds]);
    } catch (err) {
      console.error(`[graphSync] Failed to auto-archive thread ${agg.threadId}:`, err);
    }
  }

  try {
    const { getBundleRule, holdThread, getNextDeliveryTime } = await import("@/services/db/bundleRules");
    const rule = await getBundleRule(accountId, category);
    if (rule?.delivery_enabled && rule.delivery_schedule) {
      const heldUntil = getNextDeliveryTime(JSON.parse(rule.delivery_schedule));
      await holdThread(accountId, agg.threadId, category, heldUntil);
    }
  } catch (err) {
    console.error(`[graphSync] Failed to check bundle rule for thread ${agg.threadId}:`, err);
  }
}

async function loadAutoArchiveCategories(): Promise<Set<string>> {
  const raw = await getSetting("auto_archive_categories");
  if (!raw) return new Set();
  return new Set(raw.split(",").map((s) => s.trim()).filter(Boolean));
}

/** Move messages to another folder (IDs stay stable thanks to immutable IDs). */
export async function moveMessages(
  client: GraphClient,
  messageIds: string[],
  destinationFolderId: string,
): Promise<void> {
  if (messageIds.length === 0) return;
  const responses = await client.batch(
    messageIds.map((id, i) => ({
      id: String(i),
      method: "POST" as const,
      url: `/me/messages/${encodeURIComponent(id)}/move`,
      body: { destinationId: destinationFolderId },
    })),
  );
  assertBatchOk(responses);
}

// ---------------------------------------------------------------------------
// Graph reads
// ---------------------------------------------------------------------------

/**
 * Walk a folder's message delta once to obtain a delta link. Only IDs and a
 * few flags are requested, so this is cheap even for very large folders.
 */
async function establishDeltaBaseline(client: GraphClient, folderId: string): Promise<string | null> {
  return client.forEachPage<GraphMessage>(
    `/me/mailFolders/${encodeURIComponent(folderId)}/messages/delta?$select=${DELTA_SELECT}`,
    () => undefined,
    { prefer: [`odata.maxpagesize=${DELTA_PAGE_SIZE}`] },
  );
}

function folderMessagesUrl(folderId: string, since: string, query: string): string {
  const filter = encodeURIComponent(`receivedDateTime ge ${since}`);
  return `/me/mailFolders/${encodeURIComponent(folderId)}/messages?$filter=${filter}&$orderby=${encodeURIComponent("receivedDateTime desc")}&$top=${MESSAGE_PAGE_SIZE}&${query}`;
}

async function fetchConversation(client: GraphClient, conversationId: string): Promise<GraphMessage[]> {
  const filter = encodeURIComponent(`conversationId eq ${odataQuote(conversationId)}`);
  return withAttachmentFallback(() =>
    client.getAll<GraphMessage>(
      `/me/messages?$filter=${filter}&$top=${MESSAGE_PAGE_SIZE}&${messageQuery()}`,
      { prefer: [HTML_BODY_PREFERENCE] },
    ),
  );
}

// ---------------------------------------------------------------------------
// Initial sync
// ---------------------------------------------------------------------------

/** Inbox first so the most useful mail shows up early, then Sent, then the rest. */
function folderSyncOrder(a: GraphFolderLabel, b: GraphFolderLabel): number {
  const rank = (f: GraphFolderLabel): number => (f.labelId === "INBOX" ? 0 : f.labelId === "SENT" ? 1 : 2);
  return rank(a) - rank(b);
}

/**
 * First sync for a Microsoft 365 account: mirror folders and categories,
 * record a delta baseline per folder, then download the last `daysBack` days.
 */
export async function graphInitialSync(
  accountId: string,
  daysBack = 365,
  onProgress?: GraphSyncProgressCallback,
): Promise<void> {
  const client = getGraphClient(accountId);

  onProgress?.({ phase: "labels", current: 0, total: 1 });
  const context = await syncFoldersAndCategories(client, accountId);
  onProgress?.({ phase: "labels", current: 1, total: 1 });

  // Baselines are taken BEFORE downloading so nothing that arrives during the
  // download is missed — the first delta sync simply replays it.
  const state: GraphSyncState = { deltaLinks: {}, retryThreadIds: [] };
  let baselined = 0;
  onProgress?.({ phase: "threads", current: 0, total: context.folders.length });
  await parallelLimit(context.folders, FOLDER_CONCURRENCY, async (folder) => {
    const link = await establishDeltaBaseline(client, folder.folderId);
    if (link) state.deltaLinks[folder.folderId] = link;
    onProgress?.({ phase: "threads", current: ++baselined, total: context.folders.length });
  });

  const since = sinceIso(daysBack);
  const aggregates = new Map<string, ThreadAggregate>();
  const estimatedTotal = context.folders.reduce((sum, f) => sum + f.totalItemCount, 0);
  let stored = 0;

  for (const folder of [...context.folders].sort(folderSyncOrder)) {
    if (folder.totalItemCount === 0) continue;
    try {
      await withAttachmentFallback(() =>
        client.forEachPage<GraphMessage>(
          folderMessagesUrl(folder.folderId, since, messageQuery()),
          async (items) => {
            const parsed = items.map((item) => graphMessageToParsed(item, folder.labelId));
            await withTransaction(async () => {
              const touched = new Set<ThreadAggregate>();
              for (const msg of parsed) {
                // The folder is known from the request, so trust it over the payload
                msg.folderId = folder.folderId;
                touched.add(mergeIntoAggregate(aggregates, msg));
              }
              for (const agg of touched) await writeThread(accountId, agg);
              for (const msg of parsed) await writeMessage(accountId, msg);
            });
            stored += parsed.length;
            onProgress?.({ phase: "messages", current: stored, total: Math.max(stored, estimatedTotal) });
          },
          { prefer: [HTML_BODY_PREFERENCE] },
        ),
      );
    } catch (err) {
      // One bad folder shouldn't block the rest of the mailbox
      console.error(`[graphSync] Failed to sync folder ${folder.labelName}:`, err);
    }
  }

  const autoArchiveCategories = await loadAutoArchiveCategories();
  for (const agg of aggregates.values()) {
    try {
      await categorizeThread(client, accountId, agg, context, autoArchiveCategories);
    } catch (err) {
      console.error(`[graphSync] Failed to categorize thread ${agg.threadId}:`, err);
    }
  }

  await saveSyncState(accountId, state);
  await updateAccountSyncState(accountId, `graph:${Date.now()}`);
  onProgress?.({ phase: "done", current: stored, total: stored });
}

// ---------------------------------------------------------------------------
// Delta sync
// ---------------------------------------------------------------------------

interface DeltaChanges {
  /** Conversations that gained or changed messages. */
  conversationIds: Set<string>;
  /** Message IDs reported as added/updated per conversation. */
  changedMessageIds: Set<string>;
  /** Message IDs that left a folder (deleted, or moved elsewhere). */
  removedMessageIds: Set<string>;
}

function collectDeltaItems(items: GraphMessage[], changes: DeltaChanges): void {
  for (const item of items) {
    if (item["@removed"]) {
      changes.removedMessageIds.add(item.id);
    } else {
      changes.changedMessageIds.add(item.id);
      if (item.conversationId) changes.conversationIds.add(item.conversationId);
    }
  }
}

/** Baseline a folder with no usable delta link and queue its recent conversations. */
async function rebaselineFolder(
  client: GraphClient,
  folderId: string,
  since: string,
  state: GraphSyncState,
  changes: DeltaChanges,
): Promise<void> {
  const link = await establishDeltaBaseline(client, folderId);
  if (link) state.deltaLinks[folderId] = link;
  await client.forEachPage<GraphMessage>(
    folderMessagesUrl(folderId, since, "$select=conversationId"),
    (items) => collectDeltaItems(items, changes),
  );
}

async function threadIdsForMessages(accountId: string, messageIds: string[]): Promise<string[]> {
  if (messageIds.length === 0) return [];
  const db = await getDb();
  const threadIds = new Set<string>();
  const chunkSize = 200;
  for (let i = 0; i < messageIds.length; i += chunkSize) {
    const chunk = messageIds.slice(i, i + chunkSize);
    const placeholders = chunk.map((_, n) => `$${n + 2}`).join(", ");
    const rows = await db.select<{ thread_id: string }[]>(
      `SELECT DISTINCT thread_id FROM messages WHERE account_id = $1 AND id IN (${placeholders})`,
      [accountId, ...chunk],
    );
    for (const row of rows) threadIds.add(row.thread_id);
  }
  return [...threadIds];
}

async function storedMessageIds(accountId: string, threadId: string): Promise<Set<string>> {
  const db = await getDb();
  const rows = await db.select<{ id: string }[]>(
    "SELECT id FROM messages WHERE account_id = $1 AND thread_id = $2",
    [accountId, threadId],
  );
  return new Set(rows.map((r) => r.id));
}

interface RefreshSettings {
  autoArchiveCategories: Set<string>;
  mutedThreadIds: Set<string>;
  smartNotifications: boolean;
  notifyCategories: Set<string>;
  vipSenders: Set<string>;
}

/**
 * Re-download one conversation and make the local thread match the server.
 * Returns the messages that are new to the local DB.
 */
export async function refreshConversation(
  client: GraphClient,
  accountId: string,
  threadId: string,
  context: GraphFolderContext,
  settings: RefreshSettings,
): Promise<GraphParsedMessage[]> {
  // Don't overwrite threads with queued local changes — they'd flicker back
  const pendingOps = await getPendingOpsForResource(accountId, threadId);
  if (pendingOps.length > 0) return [];

  const known = await storedMessageIds(accountId, threadId);
  const remote = await fetchConversation(client, threadId);

  const parsed: GraphParsedMessage[] = [];
  for (const item of remote) {
    const labelId = item.parentFolderId ? context.labelByFolderId.get(item.parentFolderId) : undefined;
    // Messages in folders we don't sync (Outbox, Sync Issues, …) are ignored
    if (!labelId) continue;
    parsed.push(graphMessageToParsed(item, labelId));
  }

  if (parsed.length === 0) {
    // Nothing left on the server — drop the local thread
    for (const id of known) await deleteMessage(accountId, id);
    await deleteThread(accountId, threadId);
    return [];
  }

  parsed.sort((a, b) => a.date - b.date);
  const aggregates = new Map<string, ThreadAggregate>();
  for (const msg of parsed) {
    msg.threadId = threadId;
    mergeIntoAggregate(aggregates, msg);
  }
  const agg = aggregates.get(threadId)!;

  // No transaction here: conversations refresh concurrently and share one DB connection
  await writeThread(accountId, agg);
  for (const msg of parsed) await writeMessage(accountId, msg);
  for (const id of known) {
    if (!agg.messageIds.has(id)) await deleteMessage(accountId, id);
  }

  await categorizeThread(client, accountId, agg, context, settings.autoArchiveCategories);

  const isMuted = settings.mutedThreadIds.has(threadId);

  // Muted threads that resurface in the inbox go straight back to the archive
  if (isMuted && agg.inboxMessageIds.size > 0 && context.wellKnownIds.archive) {
    try {
      await moveMessages(client, [...agg.inboxMessageIds], context.wellKnownIds.archive);
      agg.labelIds.delete("INBOX");
      agg.labelIds.add("archive");
      await setThreadLabels(accountId, threadId, [...agg.labelIds]);
    } catch (err) {
      console.error(`[graphSync] Failed to auto-archive muted thread ${threadId}:`, err);
    }
  }

  const newMessages = parsed.filter((m) => !known.has(m.id));
  const newInbox = newMessages.filter((m) => m.labelIds.includes("INBOX") && !m.isRead);

  if (!isMuted) {
    for (const msg of newInbox) {
      const fromAddr = msg.fromAddress ?? undefined;
      const category = await getThreadCategory(accountId, threadId);
      if (shouldNotifyForMessage(settings.smartNotifications, settings.notifyCategories, settings.vipSenders, category, fromAddr)) {
        queueNewEmailNotification(msg.fromName ?? msg.fromAddress ?? "Unknown", msg.subject ?? "", threadId, accountId, fromAddr);
      }
    }
  }

  if (newInbox.length > 0) {
    try {
      await applyFiltersToMessages(accountId, newInbox);
    } catch (err) {
      console.error(`[graphSync] Failed to apply filters to thread ${threadId}:`, err);
    }
    import("@/services/smartLabels/smartLabelManager")
      .then(({ applySmartLabelsToMessages }) => applySmartLabelsToMessages(accountId, newInbox))
      .catch((err) => console.error("Smart label error:", err));
  }

  return newMessages;
}

/**
 * Incremental sync: ask each folder what changed since the last delta link,
 * then re-download just the conversations that were touched.
 * Returns the number of conversations refreshed.
 */
export async function graphDeltaSync(accountId: string, daysBack = 365): Promise<number> {
  const client = getGraphClient(accountId);
  const context = await syncFoldersAndCategories(client, accountId);
  const state = await loadSyncState(accountId);
  const since = sinceIso(daysBack);

  const changes: DeltaChanges = {
    conversationIds: new Set(),
    changedMessageIds: new Set(),
    removedMessageIds: new Set(),
  };

  await parallelLimit(context.folders, FOLDER_CONCURRENCY, async (folder) => {
    const link = state.deltaLinks[folder.folderId];
    try {
      if (!link) {
        // Folder created since the last sync
        await rebaselineFolder(client, folder.folderId, since, state, changes);
        return;
      }
      const next = await client.forEachPage<GraphMessage>(link, (items) => collectDeltaItems(items, changes), {
        prefer: [`odata.maxpagesize=${DELTA_PAGE_SIZE}`],
      });
      if (next) state.deltaLinks[folder.folderId] = next;
    } catch (err) {
      if (isDeltaExpired(err)) {
        console.warn(`[graphSync] Delta link expired for ${folder.labelName} — rebuilding`);
        await rebaselineFolder(client, folder.folderId, since, state, changes);
      } else {
        // Keep the old link so the next cycle retries this folder
        console.error(`[graphSync] Delta sync failed for folder ${folder.labelName}:`, err);
      }
    }
  });

  // Forget delta links for folders that no longer exist
  for (const folderId of Object.keys(state.deltaLinks)) {
    if (!context.labelByFolderId.has(folderId)) delete state.deltaLinks[folderId];
  }

  // A message that left a folder was either deleted or moved; refreshing its
  // thread settles which.
  const removedOnly = [...changes.removedMessageIds].filter((id) => !changes.changedMessageIds.has(id));
  const threadIds = new Set<string>([
    ...state.retryThreadIds,
    ...changes.conversationIds,
    ...(await threadIdsForMessages(accountId, removedOnly)),
  ]);
  // The delta links have already moved on, so remember refreshes that fail
  const failedThreadIds: string[] = [];

  if (threadIds.size > 0) {
    const smartNotifications = (await getSetting("smart_notifications")) !== "false";
    const settings: RefreshSettings = {
      autoArchiveCategories: await loadAutoArchiveCategories(),
      mutedThreadIds: await getMutedThreadIds(accountId),
      smartNotifications,
      notifyCategories: new Set(
        ((await getSetting("notify_categories")) ?? "Primary").split(",").map((s) => s.trim()).filter(Boolean),
      ),
      vipSenders: smartNotifications ? await getVipSenders(accountId) : new Set<string>(),
    };

    await parallelLimit([...threadIds], THREAD_CONCURRENCY, async (threadId) => {
      try {
        await refreshConversation(client, accountId, threadId, context, settings);
      } catch (err) {
        console.error(`[graphSync] Failed to refresh conversation ${threadId}:`, err);
        failedThreadIds.push(threadId);
      }
    });
  }

  state.retryThreadIds = failedThreadIds;
  await saveSyncState(accountId, state);
  await updateAccountSyncState(accountId, `graph:${Date.now()}`);

  if (threadIds.size > 0) {
    import("@/services/ai/categorizationManager")
      .then(({ categorizeNewThreads }) => categorizeNewThreads(accountId))
      .catch((err) => console.error("Categorization error:", err));
  }

  return threadIds.size;
}

/** Forget the stored delta links (used when an account is re-synced from scratch). */
export async function clearGraphSyncState(accountId: string): Promise<void> {
  await saveSyncState(accountId, { deltaLinks: {}, retryThreadIds: [] });
}
