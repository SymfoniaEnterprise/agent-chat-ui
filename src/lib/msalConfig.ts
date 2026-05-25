// Phase 02 (02-chat-ui-b2c-login) — locked B2C constants per CONTEXT.md D-01/D-02/D-03.
// Reused from existing AgentGateway Front app registration. New chat-ui hostnames must
// be added to that app's SPA Redirect URIs (Plan 02-00 / runbook 02-03).
//
// Token-shape contract enforced by orchestrator/auth/user_jwt.py:validate_user_token:
//   aud == NEXT_PUBLIC_ORCH_API_AUDIENCE  (c5b0049f-737a-4e4c-a302-204f2340de85)
//   iss == https://${tenant}.b2clogin.com/${tenant-id}/v2.0/
//
// Phase 01 UAT (2026-05-21) confirmed `response_type=token` (NOT id_token) returns
// an access_token with aud=API; id_token has aud=SPA which the orchestrator rejects.
// MSAL `loginRedirect({ scopes: [...] })` returns BOTH id_token AND access_token in
// the redirect; we consume tokenResponse.accessToken, never tokenResponse.idToken.

import { BrowserCacheLocation, type Configuration } from "@azure/msal-browser";

// Literal property access on process.env is REQUIRED for Next.js to inline
// NEXT_PUBLIC_* at build time. Dynamic access (process.env[varName]) is NOT
// inlined — it would return undefined at runtime in the production server
// bundle (the prior `requireEnv(name)` helper hit exactly this bug — see
// hot-fix 2026-05-25, root cause in agent-chat-ui pod logs).
function requireEnv(name: string, value: string | undefined): string {
  if (!value) {
    throw new Error(
      `${name} is not set — see chat-ui/FORK-NOTES.md for the five required NEXT_PUBLIC_* build args.`,
    );
  }
  return value;
}

const tenant = requireEnv("NEXT_PUBLIC_B2C_TENANT", process.env.NEXT_PUBLIC_B2C_TENANT); // symfoniab2ctest
const policy = requireEnv("NEXT_PUBLIC_B2C_POLICY", process.env.NEXT_PUBLIC_B2C_POLICY); // b2c_1a_signup_signin_dev
const clientId = requireEnv("NEXT_PUBLIC_B2C_CLIENT_ID", process.env.NEXT_PUBLIC_B2C_CLIENT_ID); // a78a94ef-5367-442e-8ae9-e74807a883b6
const apiAud = requireEnv("NEXT_PUBLIC_ORCH_API_AUDIENCE", process.env.NEXT_PUBLIC_ORCH_API_AUDIENCE); // c5b0049f-737a-4e4c-a302-204f2340de85

// Note (msal-browser v5 deviation, Rule 3): v5 dropped `navigateToLoginRequestUrl`
// and `storeAuthStateInCookie` from the public TS type defs. Runtime still reads
// `navigateToLoginRequestUrl` if present, but TS strict mode rejects setting it
// here. The v5 defaults are already what Phase 02 wants:
//   * post-redirect navigation: lands on redirectUri (window.location.origin)
//   * storeAuthStateInCookie: false (IE11/legacy-Edge only, not relevant)
// So we drop both keys instead of casting; no test asserts on either.
export const msalConfig: Configuration = {
  auth: {
    clientId,
    authority: `https://${tenant}.b2clogin.com/${tenant}.onmicrosoft.com/${policy}`,
    knownAuthorities: [`${tenant}.b2clogin.com`],
    // SSR-safe: window only exists in the browser. The redirect lands on
    // window.location.origin in the browser; the SSR pass uses "/" which is
    // fine because the redirect is never initiated server-side.
    redirectUri: typeof window !== "undefined" ? window.location.origin : "/",
  },
  cache: {
    cacheLocation: BrowserCacheLocation.SessionStorage, // "sessionStorage" per CONTEXT specifics
  },
};

export const loginRequest = {
  // Single API access scope — yields access_token with aud=NEXT_PUBLIC_ORCH_API_AUDIENCE.
  // Do NOT add OIDC standard scopes (id-only / profile / offline-access) here:
  // B2C policies emit refresh_token for the user-flow audience regardless, and
  // adding extra scopes can cause B2C to return an id_token with aud=SPA instead
  // of an access_token with aud=API — Phase 01 UAT 2026-05-21 root cause.
  scopes: [`https://${tenant}.onmicrosoft.com/${apiAud}/access`],
};

// Re-exported so orchestratorClient.ts can request the same scopes on silent refresh.
export const apiScope = `https://${tenant}.onmicrosoft.com/${apiAud}/access`;
