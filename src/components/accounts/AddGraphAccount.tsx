import { useState, useCallback, useEffect } from "react";
import { ArrowLeft, Building2, CheckCircle2, Loader2 } from "lucide-react";
import { Modal } from "@/components/ui/Modal";
import { TextField } from "@/components/ui/TextField";
import { getAccountByEmail, insertGraphAccount } from "@/services/db/accounts";
import { getSetting, setSetting } from "@/services/db/settings";
import { useAccountStore } from "@/stores/accountStore";
import { getMicrosoftGraphProvider } from "@/services/oauth/providers";
import { startProviderOAuthFlow } from "@/services/oauth/oauthFlow";
import { fetchGraphProfile } from "@/services/graph/graphClient";

/** Remembered between accounts so a second mailbox doesn't need the IDs retyped. */
const CLIENT_ID_SETTING = "microsoft_graph_client_id";
const TENANT_SETTING = "microsoft_graph_tenant";

interface AddGraphAccountProps {
  onClose: () => void;
  onSuccess: () => void;
  onBack: () => void;
}

type Status = "idle" | "authenticating" | "done";

export function AddGraphAccount({ onClose, onSuccess, onBack }: AddGraphAccountProps) {
  const addAccount = useAccountStore((s) => s.addAccount);
  const [clientId, setClientId] = useState("");
  const [tenant, setTenant] = useState("");
  const [status, setStatus] = useState<Status>("idle");
  const [error, setError] = useState<string | null>(null);
  const [connectedEmail, setConnectedEmail] = useState<string | null>(null);

  // Pre-fill the IDs used last time
  useEffect(() => {
    Promise.all([getSetting(CLIENT_ID_SETTING), getSetting(TENANT_SETTING)])
      .then(([savedClientId, savedTenant]) => {
        if (savedClientId) setClientId((current) => current || savedClientId);
        if (savedTenant) setTenant((current) => current || savedTenant);
      })
      .catch(() => {
        // Nothing saved yet
      });
  }, []);

  const handleSignIn = useCallback(async () => {
    const trimmedClientId = clientId.trim();
    const trimmedTenant = tenant.trim();
    if (!trimmedClientId) return;

    setStatus("authenticating");
    setError(null);

    try {
      const provider = getMicrosoftGraphProvider(trimmedTenant);
      const { tokens, userInfo } = await startProviderOAuthFlow(provider, trimmedClientId);

      if (!tokens.refresh_token) {
        throw new Error("Microsoft did not return a refresh token. Check that offline_access is allowed for this app.");
      }

      // Ask Graph for the mailbox address; this also proves the token can reach the mailbox
      const profile = await fetchGraphProfile(tokens.access_token);
      const email = profile.email ?? userInfo.email;
      const displayName = profile.name ?? userInfo.name ?? null;
      if (!email) {
        throw new Error("Could not read the mailbox address from the Microsoft sign-in.");
      }

      const existing = await getAccountByEmail(email);
      if (existing) {
        throw new Error(`${email} is already connected.`);
      }

      const id = crypto.randomUUID();
      await insertGraphAccount({
        id,
        email,
        displayName: displayName || null,
        accessToken: tokens.access_token,
        refreshToken: tokens.refresh_token,
        tokenExpiresAt: Math.floor(Date.now() / 1000) + tokens.expires_in,
        oauthClientId: trimmedClientId,
        oauthTenant: trimmedTenant || null,
      });

      await setSetting(CLIENT_ID_SETTING, trimmedClientId);
      await setSetting(TENANT_SETTING, trimmedTenant);

      addAccount({
        id,
        email,
        displayName: displayName || null,
        avatarUrl: null,
        isActive: true,
        provider: "graph",
      });

      setConnectedEmail(email);
      setStatus("done");
    } catch (err) {
      console.error("Failed to add Microsoft 365 account:", err);
      setError(err instanceof Error ? err.message : String(err));
      setStatus("idle");
    }
  }, [clientId, tenant, addAccount]);

  return (
    <Modal isOpen={true} onClose={onClose} title="Add Microsoft 365 Account" width="w-full max-w-md">
      <div className="p-4">
        {status !== "done" && (
          <div className="space-y-4">
            <div className="flex items-center gap-3 mb-4">
              <div className="w-10 h-10 rounded-full bg-accent/10 flex items-center justify-center">
                <Building2 size={20} className="text-accent" />
              </div>
              <div>
                <h3 className="text-sm font-medium text-text-primary">Microsoft 365 / Exchange Online</h3>
                <p className="text-xs text-text-tertiary">
                  Work or school mailbox and calendar, connected through Microsoft Graph
                </p>
              </div>
            </div>

            <p className="text-xs text-text-secondary">
              Your Microsoft 365 administrator registers Velo once in Microsoft Entra ID and gives you
              the two IDs below. See Help → Accounts for the setup steps.
            </p>

            <TextField
              label="Application (client) ID"
              type="text"
              value={clientId}
              onChange={(e) => setClientId(e.target.value)}
              placeholder="00000000-0000-0000-0000-000000000000"
              autoFocus
            />

            <TextField
              label="Directory (tenant) ID"
              type="text"
              value={tenant}
              onChange={(e) => setTenant(e.target.value)}
              placeholder="Tenant ID or your domain (e.g. contoso.com)"
            />

            {error && (
              <div className="bg-danger/10 border border-danger/20 rounded-lg p-3 text-sm text-danger">
                {error}
              </div>
            )}

            {status === "authenticating" && (
              <div className="text-center py-2 text-text-secondary text-sm">
                <Loader2 size={20} className="animate-spin text-accent mx-auto mb-2" />
                <div>Waiting for Microsoft sign-in...</div>
                <div className="text-xs text-text-tertiary">
                  Complete the sign-in in your browser, then return here.
                </div>
              </div>
            )}

            <div className="flex justify-between pt-2">
              <button
                onClick={onBack}
                className="flex items-center gap-1 text-sm text-text-secondary hover:text-text-primary transition-colors"
              >
                <ArrowLeft size={14} />
                Back
              </button>
              <button
                onClick={handleSignIn}
                disabled={!clientId.trim() || status === "authenticating"}
                className="px-4 py-2 text-sm font-medium text-white bg-accent hover:bg-accent-hover rounded-md transition-colors disabled:opacity-50"
              >
                Sign in with Microsoft
              </button>
            </div>
          </div>
        )}

        {status === "done" && (
          <div className="text-center py-6">
            <CheckCircle2 size={32} className="text-success mx-auto mb-3" />
            <p className="text-sm font-medium text-text-primary">Microsoft 365 account added!</p>
            <p className="text-xs text-text-tertiary mt-1">
              {connectedEmail} — mail and calendar will sync automatically. The first sync can take a while.
            </p>
            <button
              onClick={onSuccess}
              className="mt-4 px-4 py-2 text-sm font-medium text-white bg-accent hover:bg-accent-hover rounded-md transition-colors"
            >
              Done
            </button>
          </div>
        )}
      </div>
    </Modal>
  );
}
