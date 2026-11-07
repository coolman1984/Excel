import * as React from "react";
import type { Stats } from "../types";

// Status bar (plan 2.5 item 6): the plain status, filter wording, selection stats,
// zoom percentage and a 50-200 slider.

interface Props {
    status: string;
    filterInfo: string | null;
    stats: Stats | null;
    zoom: number;
    onZoom: (pct: number) => void;
}

function fmt(n: number): string {
    return n.toLocaleString("en-US", { maximumFractionDigits: 4 });
}

export function StatusBar({ status, filterInfo, stats, zoom, onZoom }: Props): React.ReactElement {
    return (
        <div className="statusBar" data-testid="statusbar">
            <span className="statusLeft">{status}</span>
            {filterInfo !== null && <span className="statusFilter">{filterInfo}</span>}
            <span className="statusSpacer" />
            {stats !== null && stats.numbers > 0 && (
                <span className="statusStats">
                    Average: {fmt(stats.average)} &nbsp; Count: {fmt(stats.count)} &nbsp; Sum: {fmt(stats.sum)}
                </span>
            )}
            {stats !== null && stats.numbers === 0 && <span className="statusStats">Count: {fmt(stats.count)}</span>}
            <span className="statusZoom">
                <input
                    type="range"
                    min={50}
                    max={200}
                    step={5}
                    value={zoom}
                    onChange={(e) => onZoom(Number(e.target.value))}
                    aria-label="Zoom"
                />
                <button type="button" className="zoomValue" onClick={() => onZoom(100)} title="Reset zoom to 100%">
                    {zoom}%
                </button>
            </span>
        </div>
    );
}
