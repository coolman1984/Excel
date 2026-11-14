// The Agent panel: the owner asks, and Claude Code CLI or Codex CLI (signed in on this computer) does the work in
// ExcelApp through its MCP tools: reading and writing the workbook, querying the data store, running commands.
import * as React from "react";
import { call, on } from "../bridge";
import { addEvent, addPrompt, ends, type Entry } from "./transcript";

type AgentName = "claude" | "codex";
const LABEL: Record<AgentName, string> = { claude: "Claude Code", codex: "Codex" };

export function AgentPanel({ onClose }: { onClose: () => void }): React.ReactElement {
    const [agent, setAgent] = React.useState<AgentName>("claude");
    const [installed, setInstalled] = React.useState<{ claude: boolean; codex: boolean } | null>(null);
    const [entries, setEntries] = React.useState<Entry[]>([]);
    const [input, setInput] = React.useState("");
    const [busy, setBusy] = React.useState(false);
    const fresh = React.useRef(true);   // the next message starts a new conversation
    const listRef = React.useRef<HTMLDivElement | null>(null);

    React.useEffect(() => {
        call<any>("agent.status").then((s) => {
            setInstalled({ claude: !!s?.claude, codex: !!s?.codex });
            if (s?.running) setBusy(true);
            if (!s?.claude && s?.codex) setAgent("codex");
        }).catch(() => setInstalled({ claude: false, codex: false }));
        return on("agent.event", (e: any) => {
            setEntries((l) => addEvent(l, e));
            if (ends(e)) setBusy(false);
        });
    }, []);

    React.useEffect(() => { listRef.current?.scrollTo({ top: listRef.current.scrollHeight }); }, [entries]);

    const send = async () => {
        const prompt = input.trim();
        if (!prompt || busy) return;
        setEntries((l) => addPrompt(l, prompt)); setInput(""); setBusy(true);
        try { await call("agent.start", { agent, prompt, newChat: fresh.current }); fresh.current = false; }
        catch (err: any) { setEntries((l) => addEvent(l, { kind: "error", text: err?.message ?? String(err) })); setBusy(false); }
    };
    const stop = () => { void call("agent.stop").catch(() => {}); };
    const newChat = () => { if (busy) return; setEntries([]); fresh.current = true; };
    const ready = installed?.[agent] ?? false;

    return (
        <aside className="agentPanel" aria-label="Agent" data-testid="agent-panel">
            <header className="agentHead">
                <strong>Agent</strong>
                <select aria-label="Agent program" value={agent} disabled={busy} onChange={(e) => { setAgent(e.target.value as AgentName); fresh.current = true; }}>
                    {(["claude", "codex"] as AgentName[]).map((a) => (
                        <option key={a} value={a}>{LABEL[a]}{installed && !installed[a] ? " (not installed)" : ""}</option>
                    ))}
                </select>
                <button type="button" onClick={newChat} disabled={busy} title="Start a new conversation">New chat</button>
                <button type="button" className="agentClose" onClick={onClose} aria-label="Close the agent panel">×</button>
            </header>
            <div className="agentList" ref={listRef} role="log" aria-live="polite">
                {entries.length === 0 && (
                    <div className="agentHint">
                        {installed && !ready
                            ? `${LABEL[agent]} is not installed on this computer. Install it and sign in once in a terminal (${agent === "claude" ? "npm install -g @anthropic-ai/claude-code, then claude" : "npm install -g @openai/codex, then codex"}).`
                            : "Ask in your own words. The agent reads and changes the open workbook and queries the data store for you, for example: \"Total sales by region on a new summary\", or \"Find rows where the cost is missing and mark them yellow\"."}
                    </div>
                )}
                {entries.map((e, i) => (
                    <div key={i} className={`agentEntry agent-${e.kind}`} dir="auto">
                        {e.kind === "tool" ? <span className="agentTool">⚙ {e.text}</span> : e.text}
                    </div>
                ))}
                {busy && <div className="agentEntry agent-working" aria-hidden>…</div>}
            </div>
            <div className="agentInput">
                <textarea dir="auto" value={input} placeholder={`Ask ${LABEL[agent]}...`} rows={3} aria-label="Message to the agent"
                    onChange={(e) => setInput(e.target.value)}
                    onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); void send(); } }} />
                {busy
                    ? <button type="button" className="agentStop" onClick={stop}>Stop</button>
                    : <button type="button" className="agentSend" onClick={() => void send()} disabled={!input.trim() || !ready}>Send</button>}
            </div>
        </aside>
    );
}
