"use client";

// Phase 02 (02-chat-ui-b2c-login) Plan 02-01b Task 3 -- empty-state gate.
//
// Per CONTEXT D-02:
//   - explicit "Zaloguj" button (no auto-redirect on app boot)
//   - redirect flow (loginRedirect) -- the popup variant is intentionally
//     not used; redirect survives popup blockers and matches the existing
//     AgentGateway SPA pattern
//   - Polish UX strings ("Zaloguj", "Zaloguj sie aby porozmawiac z agentami")
//
// W-02 closure: LoginGate listens for ACTIVE_ACCOUNT_CHANGED. When
// orchestratorClient calls msalInstance.setActiveAccount(null) on
// interaction_required, msal-browser v5 fires ACTIVE_ACCOUNT_CHANGED
// (NOTE: v5 dropped the `ACCOUNT_REMOVED` event the plan was written
// against -- see SUMMARY deviations). This handler reads the current
// account list at that moment; if empty, it flips hasAccount to false and
// the Zaloguj button re-renders. This is the only wire that converts a
// thrown `interaction_required` from a message-send into a visible
// re-login prompt. We also still handle LOGOUT_SUCCESS for explicit
// sign-out paths (future-proofing -- no logout button in Phase 02).

import { useEffect, useState, type ReactNode } from "react";
import {
  EventType,
  type AuthenticationResult,
  type EventMessage,
} from "@azure/msal-browser";
import { msalInstance, msalReady } from "@/lib/msalInstance";
import { loginRequest } from "@/lib/msalConfig";

export function LoginGate({ children }: { children: ReactNode }) {
  const [ready, setReady] = useState(false);
  const [hasAccount, setHasAccount] = useState(false);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      await msalReady();
      if (cancelled) return;
      setHasAccount(msalInstance.getAllAccounts().length > 0);
      setReady(true);
    })();

    const id = msalInstance.addEventCallback((event: EventMessage) => {
      if (
        event.eventType === EventType.LOGIN_SUCCESS ||
        event.eventType === EventType.ACQUIRE_TOKEN_SUCCESS
      ) {
        const result = event.payload as AuthenticationResult | null;
        if (result?.account) {
          msalInstance.setActiveAccount(result.account);
          setHasAccount(true);
        }
      }
      if (event.eventType === EventType.LOGOUT_SUCCESS) {
        setHasAccount(false);
      }
      if (event.eventType === EventType.ACTIVE_ACCOUNT_CHANGED) {
        // W-02 closure: ACTIVE_ACCOUNT_CHANGED arrives from
        // orchestratorClient's setActiveAccount(null) call on
        // interaction_required. msal-browser v5 emits this event without
        // a payload, so we re-read the current account list to decide.
        // If there are no accounts left active, the gate falls back to
        // empty state and the Zaloguj button re-renders.
        setHasAccount(msalInstance.getAllAccounts().length > 0);
      }
    });

    return () => {
      cancelled = true;
      if (id) msalInstance.removeEventCallback(id);
    };
  }, []);

  // SSR + initial hydration: render nothing until msalReady resolves to
  // avoid a flash of empty-state during the initialize() round-trip.
  if (!ready) return null;

  if (!hasAccount) {
    return (
      <div
        style={{
          minHeight: "60vh",
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          justifyContent: "center",
          gap: "1.5rem",
          padding: "2rem",
        }}
      >
        <p
          style={{
            fontSize: "1.1rem",
            color: "var(--foreground, #333)",
            textAlign: "center",
          }}
        >
          Zaloguj sie aby porozmawiac z agentami
        </p>
        <button
          type="button"
          onClick={() => {
            void msalInstance.loginRedirect(loginRequest);
          }}
          style={{
            padding: "0.75rem 2rem",
            fontSize: "1rem",
            fontWeight: 600,
            cursor: "pointer",
            borderRadius: "0.5rem",
            border: "1px solid var(--border, #ccc)",
            background: "var(--primary, #1f6feb)",
            color: "white",
          }}
        >
          Zaloguj
        </button>
      </div>
    );
  }

  return <>{children}</>;
}
