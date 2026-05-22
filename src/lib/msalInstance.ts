// Phase 02 (02-chat-ui-b2c-login) Plan 02-01a Task 2 — singleton + bootstrap.
// Importers MUST `await msalReady()` before calling `getAllAccounts()`/login*.
//
// MSAL.js 3.x+ requires explicit `.initialize()` before any other call, and
// `handleRedirectPromise()` MUST be awaited so a fresh return from a B2C
// redirect is processed before any account/token lookup elsewhere in the app.

import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig } from "./msalConfig";

// Singleton. msal-browser 3.x+ requires explicit .initialize() — call it
// before handleRedirectPromise() so a fresh return from a B2C redirect is
// processed before any account/token lookup elsewhere.
export const msalInstance = new PublicClientApplication(msalConfig);

let _ready: Promise<void> | null = null;

/**
 * Idempotent bootstrap. Safe to call repeatedly — the underlying promise is
 * memoised, so initialize() and handleRedirectPromise() each run exactly once.
 *
 * Callers (LoginGate, orchestratorClient) MUST await this before reading
 * accounts. On the server (SSR), this is a no-op because msal-browser is
 * browser-only.
 */
export function msalReady(): Promise<void> {
  if (_ready) return _ready;
  _ready = (async () => {
    if (typeof window === "undefined") return; // SSR no-op
    await msalInstance.initialize();
    await msalInstance.handleRedirectPromise(); // MUST be awaited per MSAL docs
  })();
  return _ready;
}
