"use client";

// Phase 02 follow-up (2026-05-25): explicit picker dla agentów HR / FK.
// Bez tego komponentu user był zaklinowany na agencie z `NEXT_PUBLIC_ASSISTANT_ID`
// build-arg (default `hr`). Selector używa istniejącego `useQueryState("assistantId")`
// z nuqs — ten sam state co StreamProvider, więc useStream natychmiast widzi zmianę.
//
// UX decyzja: zmiana agenta resetuje thread (setThreadId(null)). Powód: mid-conversation
// switch byłby konfuzujący — kolejny message poszedłby do nowego agenta, ale historia
// rozmowy byłaby pomieszana. Reset = czysty nowy thread z wybranym agentem.

import { useQueryState } from "nuqs";

type AgentOption = {
  id: string;
  label: string;
  hint: string;
};

const AGENTS: AgentOption[] = [
  { id: "hr", label: "HR", hint: "urlopy, kadry" },
  { id: "fk", label: "FK", hint: "faktury, księgowość" },
];

export function AgentSelector() {
  const envDefault = process.env.NEXT_PUBLIC_ASSISTANT_ID || "hr";
  const [assistantId, setAssistantId] = useQueryState("assistantId", {
    defaultValue: envDefault,
  });
  const [, setThreadId] = useQueryState("threadId");

  const current = assistantId || envDefault;

  const handleChange = (newId: string) => {
    if (newId === current) return;
    setAssistantId(newId);
    setThreadId(null); // reset thread — fresh conversation z nowym agentem
  };

  return (
    <label className="flex items-center gap-2 text-sm">
      <span className="text-muted-foreground">Agent:</span>
      <select
        value={current}
        onChange={(e) => handleChange(e.target.value)}
        aria-label="Wybierz agenta"
        className="rounded-md border border-input bg-background px-2 py-1 text-sm hover:bg-accent hover:text-accent-foreground focus:outline-none focus:ring-2 focus:ring-ring"
      >
        {AGENTS.map((a) => (
          <option key={a.id} value={a.id} title={a.hint}>
            {a.label} ({a.hint})
          </option>
        ))}
      </select>
    </label>
  );
}
