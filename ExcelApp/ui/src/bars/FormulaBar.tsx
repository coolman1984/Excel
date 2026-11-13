import * as React from "react";
import { FUNCTIONS, suggest, typedFunctionPrefix, activeCall, acceptSuggestion, fnInfo, type FnInfo } from "../functions";

// Formula bar (plan 2.5 item 3): the name box (C12 or C12:F40) and the formula box
// showing raw text. Enter commits, Esc cancels. As in Excel, typing a function name
// lists matching functions (Tab or a click inserts one), the argument hint follows the
// caret, and the fx button opens Insert Function.

interface Props {
    address: string;
    raw: string;
    onNameSubmit: (address: string) => void;
    onRawCommit: (text: string) => void;
    visible: boolean;
}

export function FormulaBar({ address, raw, onNameSubmit, onRawCommit, visible }: Props): React.ReactElement | null {
    const [name, setName] = React.useState(address);
    const [text, setText] = React.useState(raw);
    const [editing, setEditing] = React.useState(false);
    const [caret, setCaret] = React.useState(0);
    const [pick, setPick] = React.useState(0);
    const [insertOpen, setInsertOpen] = React.useState(false);
    const box = React.useRef<HTMLInputElement | null>(null);

    React.useEffect(() => setName(address), [address]);
    React.useEffect(() => {
        if (!editing) setText(raw);
    }, [raw, editing]);

    const prefix = editing ? typedFunctionPrefix(text, caret) : null;
    const list = prefix ? suggest(prefix) : [];
    const hint = editing ? activeCall(text, caret) : null;
    const hintInfo = hint ? fnInfo(hint.name) : undefined;

    React.useEffect(() => setPick(0), [prefix]);

    const place = (next: string, at: number) => {
        setEditing(true);
        setText(next);
        setCaret(at);
        requestAnimationFrame(() => {
            const el = box.current;
            if (el) { el.focus(); el.setSelectionRange(at, at); }
        });
    };

    const accept = (f: FnInfo) => {
        const r = acceptSuggestion(text, caret, f.name);
        place(r.text, r.caret);
    };

    if (!visible) return null;

    return (
        <div className="formulaBar" data-testid="formulabar">
            <input
                className="nameBox"
                value={name}
                aria-label="Name box"
                onChange={(e) => setName(e.target.value)}
                onKeyDown={(e) => {
                    if (e.key === "Enter") {
                        onNameSubmit(name);
                        (e.target as HTMLInputElement).blur();
                    }
                }}
            />
            <button type="button" className="fxLabel fxButton" title="Insert Function" aria-label="Insert Function" data-testid="fx-button" onClick={() => setInsertOpen(true)}>
                fx
            </button>
            <div className="formulaWrap">
                <input
                    ref={box}
                    className="formulaBox"
                    value={text}
                    aria-label="Formula box"
                    spellCheck={false}
                    onChange={(e) => {
                        setEditing(true);
                        setText(e.target.value);
                        setCaret(e.target.selectionStart ?? e.target.value.length);
                    }}
                    onSelect={(e) => setCaret((e.target as HTMLInputElement).selectionStart ?? 0)}
                    onKeyDown={(e) => {
                        if (list.length > 0 && (e.key === "ArrowDown" || e.key === "ArrowUp")) {
                            e.preventDefault();
                            setPick((p) => (p + (e.key === "ArrowDown" ? 1 : list.length - 1)) % list.length);
                            return;
                        }
                        if (list.length > 0 && e.key === "Tab") {
                            e.preventDefault();
                            accept(list[pick] ?? list[0]);
                            return;
                        }
                        if (e.key === "Enter") {
                            onRawCommit(text);
                            setEditing(false);
                            (e.target as HTMLInputElement).blur();
                        } else if (e.key === "Escape") {
                            setEditing(false);
                            setText(raw);
                            (e.target as HTMLInputElement).blur();
                        }
                    }}
                    onBlur={() => setTimeout(() => setEditing(false), 150)}
                />
                {list.length > 0 && (
                    <ul className="fnSuggest" role="listbox" data-testid="fn-suggest">
                        {list.map((f, i) => (
                            <li
                                key={f.name}
                                role="option"
                                aria-selected={i === pick}
                                className={i === pick ? "active" : ""}
                                onMouseDown={(e) => { e.preventDefault(); accept(f); }}
                                title={f.desc}
                            >
                                <span className="fnName">{f.name}</span>
                                <span className="fnDesc">{f.desc}</span>
                            </li>
                        ))}
                    </ul>
                )}
                {list.length === 0 && hintInfo && (
                    <div className="fnHint" data-testid="fn-hint">
                        <b>{hintInfo.name}</b>(
                        {hintInfo.args.split(",").map((a, i, all) => (
                            <span key={i} className={i === Math.min(hint!.arg, all.length - 1) ? "argNow" : ""}>
                                {a.trim()}
                                {i < all.length - 1 ? ", " : ""}
                            </span>
                        ))}
                        )
                    </div>
                )}
            </div>
            {insertOpen && (
                <InsertFunction
                    onClose={() => setInsertOpen(false)}
                    onPick={(f) => {
                        setInsertOpen(false);
                        const base = editing && text.startsWith("=") ? text : "=";
                        place(base + f.name + "(", (base + f.name + "(").length);
                    }}
                />
            )}
        </div>
    );
}

const CATEGORIES = ["All", "Math", "Statistical", "Text", "Logical", "Lookup", "Date & Time", "Information", "Financial"] as const;

function InsertFunction({ onClose, onPick }: { onClose: () => void; onPick: (f: FnInfo) => void }): React.ReactElement {
    const [query, setQuery] = React.useState("");
    const [cat, setCat] = React.useState<(typeof CATEGORIES)[number]>("All");
    const q = query.trim().toUpperCase();
    const shown = FUNCTIONS.filter((f) => (cat === "All" || f.cat === cat) && (q === "" || f.name.includes(q) || f.desc.toUpperCase().includes(q))).sort((a, b) => a.name.localeCompare(b.name));
    const [sel, setSel] = React.useState(0);
    React.useEffect(() => setSel(0), [query, cat]);
    const cur = shown[Math.min(sel, shown.length - 1)];
    return (
        <div className="dialogBackdrop" onClick={onClose}>
            <div className="dialog insertFn" onClick={(e) => e.stopPropagation()} data-testid="insert-function">
                <div className="dialogTitle">
                    Insert Function
                    <button type="button" className="dialogClose" onClick={onClose} aria-label="Close">&times;</button>
                </div>
                <div className="dialogBody">
                    <label className="ifRow">
                        Search for a function:
                        <input
                            autoFocus
                            value={query}
                            onChange={(e) => setQuery(e.target.value)}
                            onKeyDown={(e) => {
                                if (e.key === "ArrowDown") { e.preventDefault(); setSel((s) => Math.min(shown.length - 1, s + 1)); }
                                else if (e.key === "ArrowUp") { e.preventDefault(); setSel((s) => Math.max(0, s - 1)); }
                                else if (e.key === "Enter" && cur) onPick(cur);
                                else if (e.key === "Escape") onClose();
                            }}
                            placeholder="Type a short description or a name"
                        />
                    </label>
                    <label className="ifRow">
                        Or select a category:
                        <select value={cat} onChange={(e) => setCat(e.target.value as (typeof CATEGORIES)[number])}>
                            {CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
                        </select>
                    </label>
                    <ul className="ifList" role="listbox">
                        {shown.map((f, i) => (
                            <li key={f.name} className={i === sel ? "active" : ""} onClick={() => setSel(i)} onDoubleClick={() => onPick(f)}>
                                {f.name}
                            </li>
                        ))}
                    </ul>
                    {cur && (
                        <div className="ifDetail">
                            <b>{cur.name}({cur.args})</b>
                            <div>{cur.desc}</div>
                        </div>
                    )}
                    <div className="dialogButtons">
                        <button type="button" disabled={!cur} onClick={() => cur && onPick(cur)}>OK</button>
                        <button type="button" onClick={onClose}>Cancel</button>
                    </div>
                </div>
            </div>
        </div>
    );
}
