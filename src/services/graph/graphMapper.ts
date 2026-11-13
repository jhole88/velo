import type { ParsedMessage, ParsedAttachment } from "../gmail/messageParser";
import { parseAuthenticationResults } from "../gmail/authParser";

// ---------------------------------------------------------------------------
// Graph resource shapes (only the fields Velo reads)
// ---------------------------------------------------------------------------

export interface GraphEmailAddress {
  emailAddress?: { name?: string | null; address?: string | null };
}

export interface GraphAttachment {
  "@odata.type"?: string;
  id: string;
  name?: string | null;
  contentType?: string | null;
  size?: number | null;
  isInline?: boolean | null;
  contentId?: string | null;
}

export interface GraphMessage {
  id: string;
  conversationId?: string | null;
  parentFolderId?: string | null;
  subject?: string | null;
  bodyPreview?: string | null;
  body?: { contentType?: string; content?: string | null } | null;
  from?: GraphEmailAddress | null;
  sender?: GraphEmailAddress | null;
  toRecipients?: GraphEmailAddress[] | null;
  ccRecipients?: GraphEmailAddress[] | null;
  bccRecipients?: GraphEmailAddress[] | null;
  replyTo?: GraphEmailAddress[] | null;
  receivedDateTime?: string | null;
  sentDateTime?: string | null;
  createdDateTime?: string | null;
  isRead?: boolean | null;
  isDraft?: boolean | null;
  importance?: string | null;
  flag?: { flagStatus?: string | null } | null;
  hasAttachments?: boolean | null;
  internetMessageId?: string | null;
  internetMessageHeaders?: { name: string; value: string }[] | null;
  categories?: string[] | null;
  attachments?: GraphAttachment[] | null;
  /** Present on delta results for items that left the folder. */
  "@removed"?: { reason?: string };
}

export interface GraphMailFolder {
  id: string;
  displayName: string;
  parentFolderId?: string | null;
  totalItemCount?: number | null;
  unreadItemCount?: number | null;
  "@removed"?: { reason?: string };
}

export interface GraphCategory {
  id?: string;
  displayName: string;
  color?: string | null;
}

// ---------------------------------------------------------------------------
// Folder → label mapping
// ---------------------------------------------------------------------------

/** Well-known Exchange folders that map onto Velo's system labels. */
export const WELL_KNOWN_LABELS: Record<string, { labelId: string; labelName: string; specialUse: string | null }> = {
  inbox: { labelId: "INBOX", labelName: "Inbox", specialUse: null },
  sentitems: { labelId: "SENT", labelName: "Sent", specialUse: "\\Sent" },
  drafts: { labelId: "DRAFT", labelName: "Drafts", specialUse: "\\Drafts" },
  deleteditems: { labelId: "TRASH", labelName: "Trash", specialUse: "\\Trash" },
  junkemail: { labelId: "SPAM", labelName: "Spam", specialUse: "\\Junk" },
  archive: { labelId: "archive", labelName: "Archive", specialUse: "\\Archive" },
};

/** Well-known folders that hold no user mail worth syncing. */
export const SKIPPED_WELL_KNOWN_FOLDERS = [
  "outbox",
  "syncissues",
  "conflicts",
  "localfailures",
  "serverfailures",
  "conversationhistory",
];

export const GRAPH_FOLDER_LABEL_PREFIX = "graph-folder:";
export const GRAPH_CATEGORY_LABEL_PREFIX = "graph-cat:";

export interface GraphFolderLabel {
  /** Graph folder ID. */
  folderId: string;
  /** Velo label ID ("INBOX", "SENT", … or "graph-folder:{id}"). */
  labelId: string;
  /** Display name — full path for nested user folders. */
  labelName: string;
  type: "system" | "user";
  specialUse: string | null;
  totalItemCount: number;
  unreadItemCount: number;
}

export function folderLabelId(folderId: string): string {
  return `${GRAPH_FOLDER_LABEL_PREFIX}${folderId}`;
}

export function categoryLabelId(displayName: string): string {
  return `${GRAPH_CATEGORY_LABEL_PREFIX}${displayName}`;
}

export function isCategoryLabel(labelId: string): boolean {
  return labelId.startsWith(GRAPH_CATEGORY_LABEL_PREFIX);
}

export function categoryNameFromLabel(labelId: string): string {
  return labelId.slice(GRAPH_CATEGORY_LABEL_PREFIX.length);
}

/**
 * Turn the flat folder list into label mappings.
 *
 * `wellKnownIds` maps a well-known name (inbox, sentitems, …) to its folder
 * ID. Skipped folders and everything beneath them are left out.
 */
export function mapFoldersToLabels(
  folders: GraphMailFolder[],
  wellKnownIds: Record<string, string>,
): GraphFolderLabel[] {
  const byId = new Map(folders.map((f) => [f.id, f]));
  const wellKnownById = new Map(Object.entries(wellKnownIds).map(([name, id]) => [id, name]));
  const skippedIds = new Set(
    SKIPPED_WELL_KNOWN_FOLDERS.map((name) => wellKnownIds[name]).filter((id): id is string => !!id),
  );

  const isSkipped = (folder: GraphMailFolder): boolean => {
    let current: GraphMailFolder | undefined = folder;
    for (let depth = 0; current && depth < 50; depth++) {
      if (skippedIds.has(current.id)) return true;
      current = current.parentFolderId ? byId.get(current.parentFolderId) : undefined;
    }
    return false;
  };

  const pathOf = (folder: GraphMailFolder): string => {
    const parts: string[] = [];
    let current: GraphMailFolder | undefined = folder;
    for (let depth = 0; current && depth < 50; depth++) {
      parts.unshift(current.displayName);
      current = current.parentFolderId ? byId.get(current.parentFolderId) : undefined;
    }
    return parts.join("/");
  };

  const result: GraphFolderLabel[] = [];
  for (const folder of folders) {
    if (folder["@removed"] || isSkipped(folder)) continue;

    const wellKnownName = wellKnownById.get(folder.id);
    const system = wellKnownName ? WELL_KNOWN_LABELS[wellKnownName] : undefined;

    result.push({
      folderId: folder.id,
      labelId: system ? system.labelId : folderLabelId(folder.id),
      labelName: system ? system.labelName : pathOf(folder),
      type: system ? "system" : "user",
      specialUse: system ? system.specialUse : null,
      totalItemCount: folder.totalItemCount ?? 0,
      unreadItemCount: folder.unreadItemCount ?? 0,
    });
  }
  return result;
}

// ---------------------------------------------------------------------------
// Outlook category colours → label colours
// ---------------------------------------------------------------------------

const CATEGORY_COLORS: Record<string, string> = {
  preset0: "#e74856", // red
  preset1: "#ff8c00", // orange
  preset2: "#aa8056", // brown
  preset3: "#f2c811", // yellow
  preset4: "#47d041", // green
  preset5: "#30c6cc", // teal
  preset6: "#73aa24", // olive
  preset7: "#00bcf2", // blue
  preset8: "#8764b8", // purple
  preset9: "#f495bf", // cranberry
  preset10: "#a0aeb2", // steel
  preset11: "#004b60", // dark steel
  preset12: "#b1adab", // gray
  preset13: "#5d5a58", // dark gray
  preset14: "#000000", // black
  preset15: "#750b1c", // dark red
  preset16: "#ca5010", // dark orange
  preset17: "#ab620d", // dark brown
  preset18: "#c19c00", // dark yellow
  preset19: "#004b1c", // dark green
  preset20: "#004b50", // dark teal
  preset21: "#0b6a0b", // dark olive
  preset22: "#002050", // dark blue
  preset23: "#32145a", // dark purple
  preset24: "#5c005c", // dark cranberry
};

export function categoryColor(preset?: string | null): string | null {
  return preset ? (CATEGORY_COLORS[preset] ?? null) : null;
}

// ---------------------------------------------------------------------------
// Message mapping
// ---------------------------------------------------------------------------

function formatAddress(entry: GraphEmailAddress): string | null {
  const address = entry.emailAddress?.address?.trim();
  if (!address) return null;
  // Address lists are split on commas elsewhere in the app, so keep names comma-free.
  const name = entry.emailAddress?.name?.replace(/[,<>"]/g, " ").replace(/\s+/g, " ").trim();
  return name && name !== address ? `${name} <${address}>` : address;
}

export function formatAddressList(list?: GraphEmailAddress[] | null): string | null {
  if (!list || list.length === 0) return null;
  const formatted = list.map(formatAddress).filter((a): a is string => a !== null);
  return formatted.length > 0 ? formatted.join(", ") : null;
}

function getHeader(msg: GraphMessage, name: string): string | null {
  const lower = name.toLowerCase();
  return msg.internetMessageHeaders?.find((h) => h.name.toLowerCase() === lower)?.value ?? null;
}

function toTimestamp(msg: GraphMessage): number {
  const iso = msg.receivedDateTime ?? msg.sentDateTime ?? msg.createdDateTime;
  const ms = iso ? Date.parse(iso) : NaN;
  return Number.isFinite(ms) ? ms : Date.now();
}

function mapAttachments(msg: GraphMessage): ParsedAttachment[] {
  return (msg.attachments ?? [])
    // Reference attachments are links to cloud files — there is nothing to download.
    .filter((att) => att["@odata.type"] !== "#microsoft.graph.referenceAttachment")
    .map((att) => ({
      filename: att.name ?? "attachment",
      mimeType: att.contentType ?? "application/octet-stream",
      size: att.size ?? 0,
      gmailAttachmentId: att.id,
      contentId: att.contentId ? att.contentId.replace(/[<>]/g, "") : null,
      isInline: att.isInline ?? false,
    }));
}

/** Label IDs for one message: its folder plus state and category labels. */
export function labelsForMessage(msg: GraphMessage, folderLabel: string): string[] {
  const labels = [folderLabel];
  if (!msg.isRead) labels.push("UNREAD");
  if (msg.flag?.flagStatus === "flagged") labels.push("STARRED");
  if (msg.importance === "high") labels.push("IMPORTANT");
  if (msg.isDraft && folderLabel !== "DRAFT") labels.push("DRAFT");
  for (const category of msg.categories ?? []) {
    labels.push(categoryLabelId(category));
  }
  return labels;
}

export interface GraphParsedMessage extends ParsedMessage {
  /** Graph folder ID the message currently lives in. */
  folderId: string | null;
  messageIdHeader: string | null;
  referencesHeader: string | null;
  inReplyToHeader: string | null;
}

/**
 * Convert a Graph message into Velo's provider-neutral message shape.
 * `folderLabel` is the label ID of the folder the message lives in.
 */
export function graphMessageToParsed(msg: GraphMessage, folderLabel: string): GraphParsedMessage {
  const isHtml = msg.body?.contentType?.toLowerCase() === "html";
  const content = msg.body?.content ?? null;
  const from = msg.from ?? msg.sender;
  const fromAddress = from?.emailAddress?.address ?? null;
  const fromName = from?.emailAddress?.name ?? null;
  const date = toTimestamp(msg);
  const attachments = mapAttachments(msg);
  const authResult = msg.internetMessageHeaders
    ? parseAuthenticationResults(msg.internetMessageHeaders)
    : null;

  return {
    id: msg.id,
    // Every Exchange message has a conversation; fall back to the message itself just in case.
    threadId: msg.conversationId ?? msg.id,
    fromAddress,
    fromName: fromName && fromName !== fromAddress ? fromName : null,
    toAddresses: formatAddressList(msg.toRecipients),
    ccAddresses: formatAddressList(msg.ccRecipients),
    bccAddresses: formatAddressList(msg.bccRecipients),
    replyTo: formatAddressList(msg.replyTo),
    subject: msg.subject ?? null,
    snippet: msg.bodyPreview ?? "",
    date,
    isRead: !!msg.isRead,
    isStarred: msg.flag?.flagStatus === "flagged",
    bodyHtml: isHtml ? content : null,
    bodyText: isHtml ? null : content,
    rawSize: 0,
    internalDate: date,
    labelIds: labelsForMessage(msg, folderLabel),
    hasAttachments: !!msg.hasAttachments || attachments.some((a) => !a.isInline),
    attachments,
    listUnsubscribe: getHeader(msg, "List-Unsubscribe"),
    listUnsubscribePost: getHeader(msg, "List-Unsubscribe-Post"),
    authResults: authResult ? JSON.stringify(authResult) : null,
    folderId: msg.parentFolderId ?? null,
    messageIdHeader: msg.internetMessageId ?? null,
    referencesHeader: getHeader(msg, "References"),
    inReplyToHeader: getHeader(msg, "In-Reply-To"),
  };
}

// ---------------------------------------------------------------------------
// Encoding helpers
// ---------------------------------------------------------------------------

/** Gmail-style base64url → standard padded base64 (what Graph's MIME endpoints expect). */
export function base64UrlToBase64(value: string): string {
  const base64 = value.replace(/-/g, "+").replace(/_/g, "/");
  const padding = (4 - (base64.length % 4)) % 4;
  return base64 + "=".repeat(padding);
}

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

/** OData string literal: single quotes are doubled. */
export function odataQuote(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}
