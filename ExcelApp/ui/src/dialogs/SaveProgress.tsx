// Progress for a save through Excel. Shows the phase text and percent, and listens to the
// saving / saved / saveFailed events the host sends. Cancel is enabled only before Excel opened.
import { useEffect, useState, type CSSProperties } from "react";
import { on } from "../bridge";

const FONT = 'Aptos, "Segoe UI Variable", "Segoe UI", Calibri, sans-serif';

const overlay: CSSProperties = {
    position: "fixed", inset: 0, background: "rgba(0,0,0,0.12)", zIndex: 2000,
    display: "flex", alignItems: "center", justifyContent: "center", fontFamily: FONT,
};
const box: CSSProperties = {
    width: 380, background: "#fff", border: "1px solid #D2D0CE",
    boxShadow: "0 8px 24px rgba(0,0,0,0.3)", borderRadius: 2, overflow: "hidden",
};
const header: CSSProperties = { background: "#F3F3F3", color: "#1b1b1b", padding: "8px 14px", fontSize: 13, fontWeight: 600, borderBottom: "1px solid #E1E1E1" };
const body: CSSProperties = { padding: "16px 18px", fontSize: 13, color: "#1b1b1b" };
const track: CSSProperties = { height: 10, background: "#EDEBE9", border: "1px solid #D2D0CE", borderRadius: 2, overflow: "hidden" };
const fill: CSSProperties = { height: "100%", background: "#107C41", transition: "width 120ms linear" };
const row: CSSProperties = { display: "flex", justifyContent: "flex-end", padding: "0 16px 16px" };
const btn: CSSProperties = {
    minWidth: 82, padding: "5px 14px", fontSize: 13, border: "1px solid #8A8886",
    background: "#F3F2F1", color: "#1b1b1b", borderRadius: 2, cursor: "pointer", fontFamily: FONT,
};

function phaseText(phase: string, percent: number): string {
    switch (phase) {
        case "opening": return "Saving through Excel...";
        case "opened": return "Excel opened the file...";
        case "writing": return `Writing changes ${percent}%`;
        case "saving": return "Saving the file...";
        case "reading": return `Reading the saved file back ${percent}%`;
        case "saved": return "Saved";
        default: return `Saving... ${percent}%`;
    }
}

/** Small save dialog. Calls onDone(true) when saved and onDone(false) when it fails or is cancelled. */
export default function SaveProgress(props: { onDone: (ok: boolean) => void }) {
    const [phase, setPhase] = useState("opening");
    const [percent, setPercent] = useState(0);
    const [canCancel, setCanCancel] = useState(true);

    useEffect(() => {
        const offSaving = on("saving", (e: any) => {
            const p: string = e?.phase ?? "saving";
            setPhase(p);
            setPercent(typeof e?.percent === "number" ? e.percent : 0);
            setCanCancel(p === "opening"); // only before OPENED, per the plan
        });
        const offSaved = on("saved", () => props.onDone(true));
        const offFailed = on("saveFailed", () => props.onDone(false));
        return () => {
            offSaving();
            offSaved();
            offFailed();
        };
    }, [props.onDone]);

    const clamped = Math.max(0, Math.min(100, percent));
    return (
        <div style={overlay} role="dialog" aria-modal="true">
            <div style={box}>
                <div style={header}>Save</div>
                <div style={body}>
                    <div style={{ marginBottom: 10 }}>{phaseText(phase, percent)}</div>
                    <div style={track}>
                        <div style={{ ...fill, width: `${clamped}%` }} />
                    </div>
                </div>
                <div style={row}>
                    <button style={{ ...btn, opacity: canCancel ? 1 : 0.5 }} disabled={!canCancel} onClick={() => props.onDone(false)}>
                        Cancel
                    </button>
                </div>
            </div>
        </div>
    );
}
