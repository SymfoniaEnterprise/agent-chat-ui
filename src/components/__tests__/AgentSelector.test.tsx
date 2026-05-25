// @vitest-environment jsdom
// Phase 02 follow-up — AgentSelector behaviour tests.
// Mocks nuqs `useQueryState` to inspect what gets called when the user picks
// a different agent: (a) the new assistantId is written, (b) threadId is
// reset to null so the next message lands in a fresh conversation.

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { render, screen, fireEvent, cleanup } from "@testing-library/react";
import { AgentSelector } from "../AgentSelector";

const setAssistantIdSpy = vi.fn();
const setThreadIdSpy = vi.fn();

vi.mock("nuqs", () => ({
  useQueryState: vi.fn((key: string, _opts?: unknown) => {
    if (key === "assistantId") return ["hr", setAssistantIdSpy];
    if (key === "threadId") return ["some-thread-id", setThreadIdSpy];
    return [null, vi.fn()];
  }),
}));

describe("AgentSelector", () => {
  beforeEach(() => {
    setAssistantIdSpy.mockClear();
    setThreadIdSpy.mockClear();
  });

  afterEach(() => {
    cleanup();
  });

  it("renders both HR and FK options", () => {
    render(<AgentSelector />);
    const select = screen.getByLabelText(/wybierz agenta/i) as HTMLSelectElement;
    const options = Array.from(select.querySelectorAll("option")).map(
      (o) => o.value,
    );
    expect(options).toEqual(["hr", "fk"]);
  });

  it("on change writes new assistantId AND resets threadId to null", () => {
    render(<AgentSelector />);
    const select = screen.getByLabelText(/wybierz agenta/i);
    fireEvent.change(select, { target: { value: "fk" } });
    expect(setAssistantIdSpy).toHaveBeenCalledWith("fk");
    expect(setThreadIdSpy).toHaveBeenCalledWith(null);
  });

  it("does not reset thread when picking the same agent", () => {
    render(<AgentSelector />);
    const select = screen.getByLabelText(/wybierz agenta/i);
    fireEvent.change(select, { target: { value: "hr" } });
    expect(setAssistantIdSpy).not.toHaveBeenCalled();
    expect(setThreadIdSpy).not.toHaveBeenCalled();
  });
});
