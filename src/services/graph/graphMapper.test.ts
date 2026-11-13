import {
  base64UrlToBase64,
  bytesToBase64,
  categoryColor,
  formatAddressList,
  graphMessageToParsed,
  labelsForMessage,
  mapFoldersToLabels,
  odataQuote,
  type GraphMailFolder,
  type GraphMessage,
} from "./graphMapper";

describe("mapFoldersToLabels", () => {
  const folders: GraphMailFolder[] = [
    { id: "f-inbox", displayName: "Inbox", parentFolderId: "root", totalItemCount: 10, unreadItemCount: 2 },
    { id: "f-sent", displayName: "Sent Items", parentFolderId: "root" },
    { id: "f-jobs", displayName: "Jobs", parentFolderId: "root" },
    { id: "f-job-937", displayName: "937 View St", parentFolderId: "f-jobs" },
    { id: "f-sync", displayName: "Sync Issues", parentFolderId: "root" },
    { id: "f-conflicts", displayName: "Conflicts", parentFolderId: "f-sync" },
  ];
  const wellKnown = { inbox: "f-inbox", sentitems: "f-sent", syncissues: "f-sync" };

  it("maps well-known folders to system labels", () => {
    const labels = mapFoldersToLabels(folders, wellKnown);
    const inbox = labels.find((l) => l.folderId === "f-inbox");
    expect(inbox).toMatchObject({ labelId: "INBOX", labelName: "Inbox", type: "system", totalItemCount: 10, unreadItemCount: 2 });
    expect(labels.find((l) => l.folderId === "f-sent")).toMatchObject({ labelId: "SENT", specialUse: "\\Sent" });
  });

  it("gives user folders a prefixed label and full path name", () => {
    const labels = mapFoldersToLabels(folders, wellKnown);
    expect(labels.find((l) => l.folderId === "f-job-937")).toMatchObject({
      labelId: "graph-folder:f-job-937",
      labelName: "Jobs/937 View St",
      type: "user",
    });
  });

  it("leaves out skipped folders and their children", () => {
    const ids = mapFoldersToLabels(folders, wellKnown).map((l) => l.folderId);
    expect(ids).not.toContain("f-sync");
    expect(ids).not.toContain("f-conflicts");
  });

  it("ignores folders flagged as removed", () => {
    const labels = mapFoldersToLabels([{ id: "x", displayName: "Gone", "@removed": { reason: "deleted" } }], {});
    expect(labels).toHaveLength(0);
  });
});

describe("labelsForMessage", () => {
  it("adds state and category labels to the folder label", () => {
    const msg: GraphMessage = {
      id: "m1",
      isRead: false,
      flag: { flagStatus: "flagged" },
      importance: "high",
      categories: ["Blue category"],
    };
    expect(labelsForMessage(msg, "INBOX")).toEqual([
      "INBOX",
      "UNREAD",
      "STARRED",
      "IMPORTANT",
      "graph-cat:Blue category",
    ]);
  });

  it("returns only the folder for a plain read message", () => {
    expect(labelsForMessage({ id: "m1", isRead: true }, "archive")).toEqual(["archive"]);
  });
});

describe("formatAddressList", () => {
  it("formats names and addresses without commas in names", () => {
    expect(
      formatAddressList([
        { emailAddress: { name: "Hole, Jeremy", address: "j@example.com" } },
        { emailAddress: { address: "plain@example.com" } },
      ]),
    ).toBe("Hole Jeremy <j@example.com>, plain@example.com");
  });

  it("returns null for empty lists", () => {
    expect(formatAddressList([])).toBeNull();
    expect(formatAddressList(null)).toBeNull();
  });
});

describe("graphMessageToParsed", () => {
  const base: GraphMessage = {
    id: "msg-1",
    conversationId: "conv-1",
    parentFolderId: "f-inbox",
    subject: "RFI 12",
    bodyPreview: "Please see attached",
    body: { contentType: "html", content: "<p>Please see attached</p>" },
    from: { emailAddress: { name: "Sam", address: "sam@example.com" } },
    toRecipients: [{ emailAddress: { name: "Jeremy", address: "j@example.com" } }],
    receivedDateTime: "2026-11-01T12:00:00Z",
    isRead: true,
    hasAttachments: true,
    internetMessageId: "<abc@example.com>",
    internetMessageHeaders: [
      { name: "List-Unsubscribe", value: "<mailto:unsub@example.com>" },
      { name: "In-Reply-To", value: "<prev@example.com>" },
    ],
    attachments: [
      { "@odata.type": "#microsoft.graph.fileAttachment", id: "att-1", name: "drawing.pdf", contentType: "application/pdf", size: 1234, isInline: false },
      { "@odata.type": "#microsoft.graph.fileAttachment", id: "att-2", name: "logo.png", contentType: "image/png", size: 50, isInline: true, contentId: "<logo@cid>" },
      { "@odata.type": "#microsoft.graph.referenceAttachment", id: "att-3", name: "link" },
    ],
  };

  it("maps core fields", () => {
    const parsed = graphMessageToParsed(base, "INBOX");
    expect(parsed).toMatchObject({
      id: "msg-1",
      threadId: "conv-1",
      fromAddress: "sam@example.com",
      fromName: "Sam",
      toAddresses: "Jeremy <j@example.com>",
      subject: "RFI 12",
      snippet: "Please see attached",
      date: Date.parse("2026-11-01T12:00:00Z"),
      isRead: true,
      isStarred: false,
      bodyHtml: "<p>Please see attached</p>",
      bodyText: null,
      labelIds: ["INBOX"],
      hasAttachments: true,
      listUnsubscribe: "<mailto:unsub@example.com>",
      folderId: "f-inbox",
      messageIdHeader: "<abc@example.com>",
      inReplyToHeader: "<prev@example.com>",
    });
  });

  it("maps file attachments and drops reference attachments", () => {
    const parsed = graphMessageToParsed(base, "INBOX");
    expect(parsed.attachments).toEqual([
      { filename: "drawing.pdf", mimeType: "application/pdf", size: 1234, gmailAttachmentId: "att-1", contentId: null, isInline: false },
      { filename: "logo.png", mimeType: "image/png", size: 50, gmailAttachmentId: "att-2", contentId: "logo@cid", isInline: true },
    ]);
  });

  it("puts plain-text bodies in bodyText", () => {
    const parsed = graphMessageToParsed({ ...base, body: { contentType: "text", content: "hello" } }, "INBOX");
    expect(parsed.bodyHtml).toBeNull();
    expect(parsed.bodyText).toBe("hello");
  });

  it("falls back to the message ID when there is no conversation", () => {
    expect(graphMessageToParsed({ id: "solo" }, "INBOX").threadId).toBe("solo");
  });
});

describe("encoding helpers", () => {
  it("converts base64url to padded base64", () => {
    expect(base64UrlToBase64("a-b_cw")).toBe("a+b/cw==");
  });

  it("encodes bytes as base64", () => {
    expect(bytesToBase64(new Uint8Array([72, 105]))).toBe("SGk=");
  });

  it("escapes single quotes for OData", () => {
    expect(odataQuote("it's")).toBe("'it''s'");
  });

  it("maps category presets to colours", () => {
    expect(categoryColor("preset7")).toBe("#00bcf2");
    expect(categoryColor("none")).toBeNull();
    expect(categoryColor(null)).toBeNull();
  });
});
