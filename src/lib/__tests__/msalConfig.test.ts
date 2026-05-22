// Phase 02 (02-chat-ui-b2c-login) Plan 02-01a Task 2 — TDD RED first.
// Asserts the resolved authority/scope/cacheLocation match the values verified
// during Phase 01 operator UAT 2026-05-21 (aud=API c5b0049f..., not aud=SPA).
import { describe, test, expect, beforeEach, afterEach, vi } from "vitest";

describe("msalConfig", () => {
  const ORIGINAL_ENV = { ...process.env };

  beforeEach(() => {
    process.env.NEXT_PUBLIC_B2C_TENANT = "symfoniab2ctest";
    process.env.NEXT_PUBLIC_B2C_POLICY = "b2c_1a_signup_signin_dev";
    process.env.NEXT_PUBLIC_B2C_CLIENT_ID = "a78a94ef-5367-442e-8ae9-e74807a883b6";
    process.env.NEXT_PUBLIC_ORCH_API_AUDIENCE = "c5b0049f-737a-4e4c-a302-204f2340de85";
    vi.resetModules();
  });

  afterEach(() => {
    process.env = { ...ORIGINAL_ENV };
  });

  test("clientId is the AgentGateway Front SPA app id (D-01)", async () => {
    const { msalConfig } = await import("../msalConfig");
    expect(msalConfig.auth.clientId).toBe("a78a94ef-5367-442e-8ae9-e74807a883b6");
  });

  test("authority is the b2clogin.com URL for the dev policy", async () => {
    const { msalConfig } = await import("../msalConfig");
    expect(msalConfig.auth.authority).toBe(
      "https://symfoniab2ctest.b2clogin.com/symfoniab2ctest.onmicrosoft.com/b2c_1a_signup_signin_dev",
    );
  });

  test("knownAuthorities contains the b2clogin host (required for B2C, see MSAL docs)", async () => {
    const { msalConfig } = await import("../msalConfig");
    expect(msalConfig.auth.knownAuthorities).toEqual(["symfoniab2ctest.b2clogin.com"]);
  });

  test("cache.cacheLocation is sessionStorage (CONTEXT specifics — closes on tab close)", async () => {
    const { msalConfig } = await import("../msalConfig");
    // msalConfig.cache is typed as optional on Configuration in msal-browser v5,
    // but we always set it. Non-null-assert to satisfy strict mode.
    expect(msalConfig.cache!.cacheLocation).toBe("sessionStorage");
  });

  test("loginRequest.scopes has exactly the API access scope (no openid/profile — D-03)", async () => {
    const { loginRequest } = await import("../msalConfig");
    expect(loginRequest.scopes).toEqual([
      "https://symfoniab2ctest.onmicrosoft.com/c5b0049f-737a-4e4c-a302-204f2340de85/access",
    ]);
  });

  test("missing env var throws a clear error pointing at FORK-NOTES.md", async () => {
    delete process.env.NEXT_PUBLIC_B2C_TENANT;
    vi.resetModules();
    await expect(async () => {
      await import("../msalConfig");
    }).rejects.toThrow(/NEXT_PUBLIC_B2C_TENANT/);
    vi.resetModules();
    await expect(async () => {
      await import("../msalConfig");
    }).rejects.toThrow(/FORK-NOTES\.md/);
  });
});
