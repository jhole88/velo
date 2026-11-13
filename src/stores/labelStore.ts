import { create } from "zustand";
import { getLabelsForAccount, deleteLabel as dbDeleteLabel, updateLabelSortOrder } from "@/services/db/labels";
import { upsertLabel } from "@/services/db/labels";
import { getGmailClient } from "@/services/gmail/tokenManager";
import { getAccount } from "@/services/db/accounts";

/**
 * Microsoft 365 accounts have folders rather than labels, so label management
 * maps onto mail folders there. Outlook categories are managed in Outlook.
 */
async function isGraphAccount(accountId: string): Promise<boolean> {
  const account = await getAccount(accountId);
  return account?.provider === "graph";
}

function assertGraphFolderLabel(labelId: string): void {
  if (!labelId.startsWith("graph-folder:")) {
    throw new Error("Outlook categories and built-in folders can't be changed from here — manage them in Outlook.");
  }
}

export interface Label {
  id: string;
  accountId: string;
  name: string;
  type: string;
  colorBg: string | null;
  colorFg: string | null;
  sortOrder: number;
}

// System labels that are already shown as nav items in the sidebar
const SYSTEM_LABEL_IDS = new Set([
  "INBOX",
  "SENT",
  "DRAFT",
  "TRASH",
  "SPAM",
  "STARRED",
  "UNREAD",
  "IMPORTANT",
  "SNOOZED",
  "CHAT",
]);

const CATEGORY_PREFIX = "CATEGORY_";

export function isSystemLabel(id: string): boolean {
  return SYSTEM_LABEL_IDS.has(id) || id.startsWith(CATEGORY_PREFIX);
}

interface LabelState {
  labels: Label[];
  isLoading: boolean;
  loadLabels: (accountId: string) => Promise<void>;
  clearLabels: () => void;
  createLabel: (accountId: string, name: string, color?: { textColor: string; backgroundColor: string }) => Promise<void>;
  updateLabel: (accountId: string, labelId: string, updates: { name?: string; color?: { textColor: string; backgroundColor: string } | null }) => Promise<void>;
  deleteLabel: (accountId: string, labelId: string) => Promise<void>;
  reorderLabels: (accountId: string, labelIds: string[]) => Promise<void>;
}

export const useLabelStore = create<LabelState>((set, get) => ({
  labels: [],
  isLoading: false,

  loadLabels: async (accountId: string) => {
    set({ isLoading: true });
    try {
      const dbLabels = await getLabelsForAccount(accountId);
      const labels: Label[] = dbLabels
        .filter((l) => !isSystemLabel(l.id))
        .map((l) => ({
          id: l.id,
          accountId: l.account_id,
          name: l.name,
          type: l.type,
          colorBg: l.color_bg,
          colorFg: l.color_fg,
          sortOrder: l.sort_order,
        }));
      set({ labels, isLoading: false });
    } catch (err) {
      console.error("Failed to load labels:", err);
      set({ isLoading: false });
    }
  },

  clearLabels: () => set({ labels: [], isLoading: false }),

  createLabel: async (accountId: string, name: string, color?: { textColor: string; backgroundColor: string }) => {
    if (await isGraphAccount(accountId)) {
      const { getEmailProvider } = await import("@/services/email/providerFactory");
      const provider = await getEmailProvider(accountId);
      const folder = await provider.createFolder(name);
      await upsertLabel({
        id: folder.id,
        accountId,
        name: folder.name,
        type: "user",
        imapFolderPath: folder.id.replace(/^graph-folder:/, ""),
      });
      await get().loadLabels(accountId);
      return;
    }
    const client = await getGmailClient(accountId);
    const gmailLabel = await client.createLabel(name, color);
    await upsertLabel({
      id: gmailLabel.id,
      accountId,
      name: gmailLabel.name,
      type: gmailLabel.type,
      colorBg: gmailLabel.color?.backgroundColor ?? null,
      colorFg: gmailLabel.color?.textColor ?? null,
    });
    await get().loadLabels(accountId);
  },

  updateLabel: async (accountId: string, labelId: string, updates: { name?: string; color?: { textColor: string; backgroundColor: string } | null }) => {
    if (await isGraphAccount(accountId)) {
      assertGraphFolderLabel(labelId);
      if (updates.name) {
        const { getEmailProvider } = await import("@/services/email/providerFactory");
        const provider = await getEmailProvider(accountId);
        await provider.renameFolder(labelId, updates.name);
        await upsertLabel({ id: labelId, accountId, name: updates.name, type: "user" });
      }
      await get().loadLabels(accountId);
      return;
    }
    const client = await getGmailClient(accountId);
    const gmailLabel = await client.updateLabel(labelId, updates);
    await upsertLabel({
      id: gmailLabel.id,
      accountId,
      name: gmailLabel.name,
      type: gmailLabel.type,
      colorBg: gmailLabel.color?.backgroundColor ?? null,
      colorFg: gmailLabel.color?.textColor ?? null,
    });
    await get().loadLabels(accountId);
  },

  deleteLabel: async (accountId: string, labelId: string) => {
    if (await isGraphAccount(accountId)) {
      assertGraphFolderLabel(labelId);
      const { getEmailProvider } = await import("@/services/email/providerFactory");
      const provider = await getEmailProvider(accountId);
      // Exchange moves the folder and its mail to Deleted Items
      await provider.deleteFolder(labelId);
      await dbDeleteLabel(accountId, labelId);
      await get().loadLabels(accountId);
      return;
    }
    const client = await getGmailClient(accountId);
    await client.deleteLabel(labelId);
    await dbDeleteLabel(accountId, labelId);
    await get().loadLabels(accountId);
  },

  reorderLabels: async (accountId: string, labelIds: string[]) => {
    const labelOrders = labelIds.map((id, index) => ({ id, sortOrder: index }));
    await updateLabelSortOrder(accountId, labelOrders);
    await get().loadLabels(accountId);
  },
}));
