// The Agent panel's conversation as a list of entries; the host sends "agent.event" {kind, text}.
export type AgentKind = "session" | "text" | "tool" | "toolError" | "status" | "error" | "done";
export interface Entry { role: "user" | "agent"; kind: AgentKind | "prompt"; text: string }

/** Adds one host event. A repeated status replaces the previous one, so "working..." lines do not pile up. */
export function addEvent(list: Entry[], e: { kind?: string; text?: string }): Entry[] {
    const kind = (e.kind ?? "status") as AgentKind;
    const text = e.text ?? "";
    if (kind === "session" || (!text && kind !== "done")) return list;
    const last = list[list.length - 1];
    if (kind === "status" && last?.kind === "status") return [...list.slice(0, -1), { role: "agent", kind, text }];
    return [...list, { role: "agent", kind, text }];
}

export function addPrompt(list: Entry[], text: string): Entry[] {
    return [...list, { role: "user", kind: "prompt", text }];
}

/** True when the event ends the agent's turn. */
export function ends(e: { kind?: string }): boolean {
    return e.kind === "done" || e.kind === "error";
}
