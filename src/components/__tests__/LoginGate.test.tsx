// @vitest-environment jsdom
// Phase 02 (02-chat-ui-b2c-login) Plan 02-01b Task 3 -- TDD GREEN.
// Behavioural tests for the LoginGate empty-state gate:
//   6: no accounts -> renders Polish empty-state hint + "Zaloguj" button
//   7: clicking "Zaloguj" calls msalInstance.loginRedirect(loginRequest)
//   8: at least one account -> renders {children}, NOT the empty state
//  10 (W-02): ACTIVE_ACCOUNT_CHANGED event (msal-browser v5 replacement for
//             v3's ACCOUNT_REMOVED -- see SUMMARY deviations) + zero accounts
//             flips hasAccount false -> empty state re-renders
import { describe, test, expect, beforeEach, afterEach, vi } from "vitest";
import { EventType, type EventMessage } from "@azure/msal-browser";
import { render, screen, fireEvent, act, cleanup } from "@testing-library/react";

const { getAllAccountsSpy, loginRedirectSpy, setActiveAccountSpy, addEventCallbackSpy, removeEventCallbackSpy, msalReadySpy } = vi.hoisted(() => ({
  getAllAccountsSpy: vi.fn(),
  loginRedirectSpy: vi.fn(() => Promise.resolve()),
  setActiveAccountSpy: vi.fn(),
  addEventCallbackSpy: vi.fn(),
  removeEventCallbackSpy: vi.fn(),
  msalReadySpy: vi.fn(() => Promise.resolve()),
}));

vi.mock("@/lib/msalInstance", () => ({
  msalInstance: {
    getAllAccounts: getAllAccountsSpy,
    loginRedirect: loginRedirectSpy,
    setActiveAccount: setActiveAccountSpy,
    addEventCallback: addEventCallbackSpy,
    removeEventCallback: removeEventCallbackSpy,
  },
  msalReady: msalReadySpy,
}));

vi.mock("@/lib/msalConfig", () => ({
  loginRequest: { scopes: ["https://symfoniab2ctest.onmicrosoft.com/c5b0049f-737a-4e4c-a302-204f2340de85/access"] },
}));

import { LoginGate } from "../LoginGate";
import { loginRequest } from "@/lib/msalConfig";

describe("LoginGate", () => {
  let registeredCallback: ((evt: EventMessage) => void) | null = null;

  beforeEach(() => {
    registeredCallback = null;
    getAllAccountsSpy.mockReset();
    loginRedirectSpy.mockClear();
    setActiveAccountSpy.mockReset();
    addEventCallbackSpy.mockReset().mockImplementation((cb: (evt: EventMessage) => void) => {
      registeredCallback = cb;
      return "callback-id";
    });
    removeEventCallbackSpy.mockReset();
    msalReadySpy.mockClear();
  });

  afterEach(() => {
    cleanup();
  });

  test("Test 6: no accounts -> renders Polish hint + 'Zaloguj' button", async () => {
    getAllAccountsSpy.mockReturnValue([]);
    render(<LoginGate><div>chat surface</div></LoginGate>);
    expect(await screen.findByText("Zaloguj sie aby porozmawiac z agentami")).toBeTruthy();
    expect(await screen.findByRole("button", { name: "Zaloguj" })).toBeTruthy();
    expect(screen.queryByText("chat surface")).toBeNull();
  });

  test("Test 7: clicking 'Zaloguj' calls msalInstance.loginRedirect(loginRequest) once", async () => {
    getAllAccountsSpy.mockReturnValue([]);
    render(<LoginGate><div>chat surface</div></LoginGate>);
    const btn = await screen.findByRole("button", { name: "Zaloguj" });
    fireEvent.click(btn);
    expect(loginRedirectSpy).toHaveBeenCalledTimes(1);
    expect(loginRedirectSpy).toHaveBeenCalledWith(loginRequest);
  });

  test("Test 8: with an account, renders children and NOT the Zaloguj button", async () => {
    getAllAccountsSpy.mockReturnValue([{ homeAccountId: "x" }]);
    render(<LoginGate><div>chat surface</div></LoginGate>);
    expect(await screen.findByText("chat surface")).toBeTruthy();
    expect(screen.queryByRole("button", { name: "Zaloguj" })).toBeNull();
    expect(screen.queryByText("Zaloguj sie aby porozmawiac z agentami")).toBeNull();
  });

  test("Test 10 (W-02): ACTIVE_ACCOUNT_CHANGED + zero accounts re-renders empty state with Zaloguj", async () => {
    // Mount with one account -> chat surface renders.
    getAllAccountsSpy.mockReturnValue([{ homeAccountId: "x" }]);
    render(<LoginGate><div>chat surface</div></LoginGate>);

    // Initially chat surface shows (one account).
    expect(await screen.findByText("chat surface")).toBeTruthy();
    expect(screen.queryByRole("button", { name: "Zaloguj" })).toBeNull();

    // The component must have registered an event callback during mount.
    expect(registeredCallback).not.toBeNull();

    // Simulate the real W-02 path: orchestratorClient calls
    // msalInstance.setActiveAccount(null) -> msal-browser v5 fires
    // ACTIVE_ACCOUNT_CHANGED -> on re-read getAllAccounts() returns [].
    getAllAccountsSpy.mockReturnValue([]);
    act(() => {
      registeredCallback!({
        eventType: EventType.ACTIVE_ACCOUNT_CHANGED,
        interactionType: null,
        payload: null,
        error: null,
        correlationId: "test-correlation-id",
        timestamp: Date.now(),
      });
    });

    // The Zaloguj button is now visible; chat surface is gone.
    expect(await screen.findByRole("button", { name: "Zaloguj" })).toBeTruthy();
    expect(screen.queryByText("chat surface")).toBeNull();
  });
});
