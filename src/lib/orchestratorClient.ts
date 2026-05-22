// Phase 02 (02-chat-ui-b2c-login) Plan 02-01b Task 3 -- fetch wrapper that
// attaches Authorization: Bearer <access_token> on every orchestrator call.
//
// Per CONTEXT D-04:
//   - silent refresh via acquireTokenSilent before each request
//   - InteractionRequiredAuthError surfaces to the caller as
//     Error("interaction_required"); the UI re-renders the Zaloguj button
//   - pre-flight `exp` decode rejects expired cached tokens (mirror of
//     tests/smoke-e2e.ps1 Test-TokenExpiry -- 30s safety margin). Avoids the
//     "expired token returns 401" UX trap operator hit on 2026-05-21.
//
// W-02 wiring: BEFORE throwing `Error("interaction_required")`, we call
// msalInstance.setActiveAccount(null). That call fires MSAL's
// ACCOUNT_REMOVED event, which LoginGate's event-callback handler turns
// into setHasAccount(false) -- re-rendering the Zaloguj button.
// Without setActiveAccount(null) the throw bubbles to the caller but the UI
// never gets the signal to switch back to empty state.
//
// Per CONTEXT D-05: NEXT_PUBLIC_ORCH_URL points at NGINX ingress, NOT the
// gateway ELB. The gateway strips Authorization
// (crates/agentgateway/src/http/jwt.rs:501, Phase 01 UAT 2026-05-21).

import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { msalInstance, msalReady } from "./msalInstance";
import { loginRequest } from "./msalConfig";

function requireEnv(name: string): string {
  const v = process.env[name];
  if (!v) {
    throw new Error(
      `${name} is not set -- see chat-ui/FORK-NOTES.md for the five required NEXT_PUBLIC_* build args.`,
    );
  }
  return v;
}

/** Decode the JWT payload section. Throws on malformed input. */
function decodeJwtPayload(token: string): { exp?: number; aud?: string; sub?: string } {
  const parts = token.split(".");
  if (parts.length !== 3 || !parts[1]) throw new Error("malformed JWT");
  const b64 = parts[1].replace(/-/g, "+").replace(/_/g, "/");
  const padded = b64 + "=".repeat((4 - (b64.length % 4)) % 4);
  // atob is available in browsers and in Node 16+; vitest's node env also has it.
  return JSON.parse(atob(padded));
}

/**
 * Token is alive iff `exp` is more than 30 seconds in the future.
 * Mirror of tests/smoke-e2e.ps1 `Test-TokenExpiry` (same 30s safety margin).
 * Returns false on any parse failure so callers route through the
 * interaction_required path rather than send a malformed token.
 */
function isTokenAlive(token: string): boolean {
  try {
    return (decodeJwtPayload(token).exp ?? 0) * 1000 > Date.now() + 30_000;
  } catch {
    return false;
  }
}

/**
 * W-02 helper: clear the active account so MSAL fires ACCOUNT_REMOVED,
 * then throw the standard interaction_required error. LoginGate's
 * event-callback observes the event and re-renders the empty state.
 * Returns `never` -- callers can `bounceToReLogin();` without a return.
 */
function bounceToReLogin(): never {
  msalInstance.setActiveAccount(null);
  throw new Error("interaction_required");
}

export type ChatMessage = { role: "user" | "assistant"; content: string };
export type ChatModel = "hr" | "fk" | "auto";
export type ChatRequestBody = { model: ChatModel; messages: ChatMessage[] };

/**
 * POST a single user-message turn to the orchestrator's chat-completions
 * endpoint. Returns the raw Response so callers can decide on streaming /
 * JSON parsing.
 *
 * Throws:
 *   - Error("not_logged_in") if MSAL has no accounts at all
 *   - Error("interaction_required") if silent refresh fails OR the cached
 *     token is expired/near-expiry/malformed (W-02 bounce path)
 *   - other Errors for unexpected msal/network failures
 */
export async function chatComplete(model: ChatModel, content: string): Promise<Response> {
  await msalReady();
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) throw new Error("not_logged_in");
  const account = accounts[0];

  let tokenResp: { accessToken: string };
  try {
    tokenResp = await msalInstance.acquireTokenSilent({ ...loginRequest, account });
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) {
      // W-02: clear active account FIRST so LoginGate's event handler
      // re-renders the empty state, THEN throw.
      bounceToReLogin();
    }
    throw e;
  }

  // Pre-flight expiry decode: protects against MSAL returning a "fresh"
  // cached token that is actually within the safety margin of expiry, or
  // returning a malformed token. Same W-02 bounce path so the UI surfaces a
  // Zaloguj prompt rather than a silent 401.
  if (!isTokenAlive(tokenResp.accessToken)) {
    bounceToReLogin();
  }

  const ORCH_URL = requireEnv("NEXT_PUBLIC_ORCH_URL");
  const body: ChatRequestBody = { model, messages: [{ role: "user", content }] };
  return fetch(`${ORCH_URL}/v1/chat/completions`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${tokenResp.accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });
}

// Exported for tests; intentionally not part of the public API surface.
export const __internal = { decodeJwtPayload, isTokenAlive, bounceToReLogin };
