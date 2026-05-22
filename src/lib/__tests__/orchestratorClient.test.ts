// Phase 02 (02-chat-ui-b2c-login) Plan 02-01b Task 3 -- TDD RED first.
// Behavioural tests for orchestratorClient.chatComplete:
//   1: no MSAL account -> Error("not_logged_in"), no fetch
//   2: happy path -> POST with Authorization: Bearer <token> + correct body
//   3: InteractionRequiredAuthError -> Error("interaction_required"), no fetch
//   4: expired cached token -> Error("interaction_required"), no fetch
//   5: malformed JWT -> Error("interaction_required"), no fetch
//   9 (W-02): InteractionRequiredAuthError path calls setActiveAccount(null) BEFORE throw
//   4b (W-02): expired-token path calls setActiveAccount(null) BEFORE throw

import { describe, test, expect, beforeEach, afterEach, vi } from "vitest";
import { InteractionRequiredAuthError } from "@azure/msal-browser";

// Hoisted spies so vi.mock factory + tests share the same fn instances.
const { getAllAccountsSpy, acquireTokenSilentSpy, setActiveAccountSpy, msalReadySpy } = vi.hoisted(() => ({
  getAllAccountsSpy: vi.fn(),
  acquireTokenSilentSpy: vi.fn(),
  setActiveAccountSpy: vi.fn(),
  msalReadySpy: vi.fn(() => Promise.resolve()),
}));

vi.mock("../msalInstance", () => ({
  msalInstance: {
    getAllAccounts: getAllAccountsSpy,
    acquireTokenSilent: acquireTokenSilentSpy,
    setActiveAccount: setActiveAccountSpy,
  },
  msalReady: msalReadySpy,
}));

vi.mock("../msalConfig", () => ({
  loginRequest: { scopes: ["https://symfoniab2ctest.onmicrosoft.com/c5b0049f-737a-4e4c-a302-204f2340de85/access"] },
  apiScope: "https://symfoniab2ctest.onmicrosoft.com/c5b0049f-737a-4e4c-a302-204f2340de85/access",
}));

/** Build a JWT with a given exp (seconds since epoch). Header + signature are stubs. */
function buildJwt(exp: number, aud = "c5b0049f-737a-4e4c-a302-204f2340de85"): string {
  const header = Buffer.from(JSON.stringify({ alg: "RS256", typ: "JWT" })).toString("base64url");
  const payload = Buffer.from(JSON.stringify({ exp, aud })).toString("base64url");
  return `${header}.${payload}.stub-sig`;
}

describe("orchestratorClient.chatComplete", () => {
  const ORIGINAL_ENV = { ...process.env };
  let fetchSpy: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    process.env.NEXT_PUBLIC_ORCH_URL = "http://ingress.example.com";
    getAllAccountsSpy.mockReset();
    acquireTokenSilentSpy.mockReset();
    setActiveAccountSpy.mockReset();
    msalReadySpy.mockClear();
    fetchSpy = vi.fn().mockResolvedValue(new Response("{}", { status: 200 }));
    vi.stubGlobal("fetch", fetchSpy);
    vi.resetModules();
  });

  afterEach(() => {
    process.env = { ...ORIGINAL_ENV };
    vi.unstubAllGlobals();
  });

  test("Test 1: no MSAL account -> Error('not_logged_in') and no fetch", async () => {
    getAllAccountsSpy.mockReturnValue([]);
    const { chatComplete } = await import("../orchestratorClient");
    await expect(chatComplete("hr", "ile mam urlopu?")).rejects.toThrow("not_logged_in");
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  test("Test 2: happy path -> POST with Authorization: Bearer + correct body", async () => {
    const futureExp = Math.floor(Date.now() / 1000) + 3600; // 1h from now
    const token = buildJwt(futureExp);
    getAllAccountsSpy.mockReturnValue([{ homeAccountId: "x" }]);
    acquireTokenSilentSpy.mockResolvedValue({ accessToken: token });

    const { chatComplete } = await import("../orchestratorClient");
    await chatComplete("hr", "ile mam urlopu?");

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    const [url, init] = fetchSpy.mock.calls[0] as [string, RequestInit];
    expect(url).toBe("http://ingress.example.com/v1/chat/completions");
    expect(init.method).toBe("POST");
    const headers = init.headers as Record<string, string>;
    expect(headers["Authorization"]).toBe(`Bearer ${token}`);
    expect(headers["Content-Type"]).toBe("application/json");
    const body = JSON.parse(init.body as string);
    expect(body).toEqual({
      model: "hr",
      messages: [{ role: "user", content: "ile mam urlopu?" }],
    });
  });

  test("Test 3: InteractionRequiredAuthError -> Error('interaction_required'), no fetch", async () => {
    getAllAccountsSpy.mockReturnValue([{ homeAccountId: "x" }]);
    acquireTokenSilentSpy.mockRejectedValue(
      new InteractionRequiredAuthError("interaction_required_error", "interaction required"),
    );

    const { chatComplete } = await import("../orchestratorClient");
    await expect(chatComplete("hr", "x")).rejects.toThrow("interaction_required");
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  test("Test 4: expired cached token (exp in past) -> Error('interaction_required'), no fetch", async () => {
    const pastExp = Math.floor(Date.now() / 1000) - 60; // 60s ago
    getAllAccountsSpy.mockReturnValue([{ homeAccountId: "x" }]);
    acquireTokenSilentSpy.mockResolvedValue({ accessToken: buildJwt(pastExp) });

    const { chatComplete } = await import("../orchestratorClient");
    await expect(chatComplete("hr", "x")).rejects.toThrow("interaction_required");
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  test("Test 4b: near-expiry token (within 30s of exp) -> Error('interaction_required'), no fetch", async () => {
    // exp is 10 seconds from now -- within the 30s safety margin -> treat as expired.
    const nearExp = Math.floor(Date.now() / 1000) + 10;
    getAllAccountsSpy.mockReturnValue([{ homeAccountId: "x" }]);
    acquireTokenSilentSpy.mockResolvedValue({ accessToken: buildJwt(nearExp) });

    const { chatComplete } = await import("../orchestratorClient");
    await expect(chatComplete("hr", "x")).rejects.toThrow("interaction_required");
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  test("Test 5: malformed JWT (no dots) -> Error('interaction_required'), no fetch", async () => {
    getAllAccountsSpy.mockReturnValue([{ homeAccountId: "x" }]);
    acquireTokenSilentSpy.mockResolvedValue({ accessToken: "no-dots-here" });

    const { chatComplete } = await import("../orchestratorClient");
    await expect(chatComplete("hr", "x")).rejects.toThrow("interaction_required");
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  test("Test 9 (W-02): on InteractionRequiredAuthError, setActiveAccount(null) called BEFORE throw", async () => {
    getAllAccountsSpy.mockReturnValue([{ homeAccountId: "x" }]);
    acquireTokenSilentSpy.mockRejectedValue(
      new InteractionRequiredAuthError("interaction_required_error", "interaction required"),
    );

    // Track the call ordering: setActiveAccount is observable BEFORE the catch sees the throw.
    let observedSetActiveAccountCallsAtThrowTime = -1;
    const { chatComplete } = await import("../orchestratorClient");

    let thrownError: Error | undefined;
    try {
      await chatComplete("hr", "x");
    } catch (e) {
      observedSetActiveAccountCallsAtThrowTime = setActiveAccountSpy.mock.calls.length;
      thrownError = e as Error;
    }

    // At the moment the throw is observed, setActiveAccount(null) MUST already have been called.
    expect(observedSetActiveAccountCallsAtThrowTime).toBe(1);
    expect(setActiveAccountSpy).toHaveBeenCalledTimes(1);
    expect(setActiveAccountSpy).toHaveBeenCalledWith(null);
    expect(thrownError?.message).toBe("interaction_required");
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  test("Test 9b (W-02): expired-cached-token path also calls setActiveAccount(null) before throw", async () => {
    const pastExp = Math.floor(Date.now() / 1000) - 60;
    getAllAccountsSpy.mockReturnValue([{ homeAccountId: "x" }]);
    acquireTokenSilentSpy.mockResolvedValue({ accessToken: buildJwt(pastExp) });

    const { chatComplete } = await import("../orchestratorClient");
    let observedAtThrowTime = -1;
    let thrownError: Error | undefined;
    try {
      await chatComplete("hr", "x");
    } catch (e) {
      observedAtThrowTime = setActiveAccountSpy.mock.calls.length;
      thrownError = e as Error;
    }

    expect(observedAtThrowTime).toBe(1);
    expect(setActiveAccountSpy).toHaveBeenCalledWith(null);
    expect(thrownError?.message).toBe("interaction_required");
    expect(fetchSpy).not.toHaveBeenCalled();
  });
});
