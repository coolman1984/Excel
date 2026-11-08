import * as React from "react";
import { call } from "../bridge";
import type { QuerySpec, ResultInfo, TableInfo } from "./contracts";
import { ResultGrid } from "./ResultGrid";
import { SpecBuilder } from "./SpecBuilder";

export function DataPage({ initial, onSpec }: { initial?: QuerySpec; onSpec: (spec: QuerySpec) => void }): React.ReactElement {
    const [tables, setTables] = React.useState<TableInfo[]>([]);
    const [projects, setProjects] = React.useState<any[]>([]);
    const [spec, setSpec] = React.useState<QuerySpec | undefined>(initial);
    const [result, setResult] = React.useState<ResultInfo>();
    const [trail, setTrail] = React.useState<ResultInfo[]>([]);
    const [error, setError] = React.useState("");
    const refresh = () => { void call<any>("data.tables.list").then(x => setTables(x.tables)); void call<any>("data.workspace.projects").then(x => setProjects(x.projects)); };
    React.useEffect(refresh, []);
    const run = React.useCallback((next: QuerySpec) => { setSpec(next); onSpec(next); setError(""); void call<ResultInfo>("data.run", { spec: next }).then(r => { setResult(r); setTrail([]); }).catch(e => setError(String(e.message ?? e))); }, [onSpec]);
    React.useEffect(() => { if (initial) run(initial); }, [initial]);
    React.useEffect(() => { if (!initial && !spec) { const first = tables.find(t => t.kind === "report"); if (first) run({ source: { kind: first.kind, id: first.id } }); } }, [tables, initial, spec, run]);
    const selected = tables.find(t => t.id === spec?.source.id) ?? projects.flatMap(p => p.tables ?? []).find((t: TableInfo) => t.id === spec?.source.id);
    const action = async (kind: string) => {
        if (!spec || !result) return;
        try {
            if (kind === "export") { const target = window.prompt("Export path (.xlsx)"); if (target) await call("data.export", { resultId: result.resultId, target }); }
            if (kind === "view") { const name = window.prompt("View name"); if (name) await call("data.views.save", { view: { id: `view-${Date.now()}`, name, spec, exception: false, folder: "" } }); }
            if (kind === "alert") { const name = window.prompt("Alert name"); if (name) await call("data.alerts.save", { alert: { id: `alert-${Date.now()}`, name, spec, measure: "rows", op: "gt", value: 0, reportTypes: [], enabled: true } }); }
            if (kind === "pin") { const data: any = await call("data.dashboard.list"); const dashboard = data.dashboards[0] ?? { id: `dashboard-${Date.now()}`, name: "Overview", tiles: [] }; dashboard.tiles = [...dashboard.tiles, { id: String(Date.now()), title: selected?.name ?? "Result", spec, measure: "rows", format: "number" }]; await call("data.dashboard.save", { dashboard }); }
        } catch (e: any) { setError(e.message); }
    };
    return <div className="data-workspace"><aside className="data-panel"><h2>Projects</h2><button onClick={async () => { const name = window.prompt("Project name"); if (name) { await call("data.workspace.create", { name }); refresh(); } }}>New project</button>{projects.map(p => <div className="project-section" key={p.id}><h3>{p.name}</h3><button onClick={async () => { const picked: any = await call("pickFile"); if (picked.path) await call("data.workspace.add", { projectId: p.id, path: picked.path }); }}>Add file</button> <button onClick={() => void call("data.workspace.refresh", { projectId: p.id })}>Refresh</button>{p.tables?.map((t: TableInfo) => <button className="table-link" key={t.id} onClick={() => run({ source: { kind: t.kind, id: t.id } })}>{t.name}</button>)}</div>)}<h2>Tables and reports</h2>{tables.map(t => <button className="table-link" key={t.id} onClick={() => run({ source: { kind: t.kind, id: t.id } })}>{t.name} <small>{t.rows.toLocaleString()} rows</small></button>)}</aside><section className="data-work"><h2>{selected?.name ?? "Choose a table"}</h2>{error && <p role="alert" className="error">{error}</p>}{spec && <><div className="inline-controls"><label>Truth <select value={spec.source.truth ?? "latest"} onChange={e => run({ ...spec, source: { ...spec.source, truth: e.target.value as "latest" | "locked" } })}><option value="latest">Latest</option><option value="locked">Locked</option></select></label><label>As of <input type="date" value={spec.source.asOf ?? ""} onChange={e => run({ ...spec, source: { ...spec.source, asOf: e.target.value || undefined } })} /></label><label>Version <input type="number" min="1" value={spec.source.version ?? ""} onChange={e => run({ ...spec, source: { ...spec.source, version: e.target.value ? Number(e.target.value) : undefined } })} /></label></div><SpecBuilder spec={spec} columns={selected?.columns ?? result?.columns ?? []} tables={tables} onRun={run} error={error} /></>}{trail.map((r, i) => <button key={r.resultId} onClick={() => { setResult(r); setTrail(trail.slice(0, i)); }}>Result {i + 1} ›</button>)}{result ? <ResultGrid result={result} onSort={col => spec && run({ ...spec, sort: [{ col, desc: spec.sort?.[0]?.col === col && !spec.sort[0].desc }] })} onDrill={r => { setTrail([...trail, result]); setResult(r); }} onAction={action} /> : <p>Select a table to explore its rows.</p>}</section></div>;
}




