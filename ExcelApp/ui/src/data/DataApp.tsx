import * as React from "react";
import { call, on } from "../bridge";
import type { QuerySpec, Question } from "./contracts";
import { DataPage } from "./DataPage";
import { IngestionPage } from "./IngestionPage";
import { AlertsPage, HistoryPage, ViewsPage } from "./ManagementPages";
import { HomePage, LibraryPage, RecipesPage, SandboxPage, SettingsPage } from "./OtherPages";
import { DashboardDesigner } from "./DashboardDesigner";
import "../styles/data.css";

const pages = ["Home", "Dashboards", "Data", "Views", "Ingestion", "History", "Alerts", "Recipes", "Sandbox", "Library", "Settings"];
const examples = ["top 10 product by sales", "bottom 5 product by sales", "total sales by region", "count by region", "average sales by region", "max sales", "min sales", "rows where sales > 100", "region is North", "show sales", "duplicates in product", "blank product", "sort by sales desc", "sales per region in October", "trend of sales by month", "compare sales and budget by product", "what changed in October", "how many rows in sales", "last 7 days of sales", "sales this month vs last month"];
export function toastReducer(state: string[], message: string): string[] { return [message, ...state].slice(0, 5); }
export function DataApp(): React.ReactElement {
    const requested = new URLSearchParams(window.location.search).get("page")?.toLowerCase();
    const [page, setPage] = React.useState(() => pages.find(p => p.toLowerCase() === requested) ?? "Home");
    const [question, setQuestion] = React.useState(""); const [understood, setUnderstood] = React.useState<Question>(); const [notices, notify] = React.useReducer(toastReducer, []); const [showNotices, setShowNotices] = React.useState(false); const [unread, setUnread] = React.useState(0); const [status, setStatus] = React.useState("Checking ingestion…"); const [spec, setSpec] = React.useState<QuerySpec>(); const [exceptions, setExceptions] = React.useState(0); const [error, setError] = React.useState("");
    React.useEffect(() => {
        let active = true;
        const refresh = () => { void call<any>("data.ingest.status").then(s => { if (active) setStatus(s.running ? "Ingestion running" : "Ingestion ready"); }).catch(() => { if (active) setStatus("Ingestion unavailable"); }); void call<any>("data.views.list").then(s => { if (active) setExceptions(s.views.filter((v: any) => v.exception).reduce((a: number, v: any) => a + (v.count ?? 0), 0)); }).catch(() => {}); };
        refresh(); const timer = window.setInterval(refresh, 10000);
        const off = ["data.notice", "data.alert", "data.loaded", "data.quarantine", "data.recipeRun"].map(event => on(event, e => { setUnread(n => n + 1); notify(e.message ?? e.alert?.message ?? (event === "data.loaded" ? "Data updated" : event === "data.quarantine" ? "A file needs your review" : "Update")); refresh(); }));
        off.push(on("data.ingest", e => setStatus(e.message ?? e.stage)));
        return () => { active = false; window.clearInterval(timer); off.forEach(f => f()); };
    }, []);
    const open = React.useCallback((s: QuerySpec) => { setSpec(s); setPage("Data"); }, []);
    React.useEffect(() => { if (!question.trim()) { setUnderstood(undefined); return; } let current = true; const timer = window.setTimeout(() => { void call<Question>("data.ask", { text: question }).then(q => { if (current) setUnderstood(q); }).catch(e => { if (current) setError(e.message); }); }, 300); return () => { current = false; window.clearTimeout(timer); }; }, [question]);
    const ask = async () => { try { const q = await call<Question>("data.ask", { text: question }); setUnderstood(q); setError(""); if (q.ok && q.spec) open(q.spec); } catch (e: any) { setError(e.message); } };
    return <div className="data-app"><aside className="data-rail"><div className="data-brand">▦ <strong>Data</strong></div><nav aria-label="Data navigation">{pages.map(p => <button type="button" className={page === p ? "active" : ""} key={p} onClick={() => setPage(p)}>{p}{p === "Views" && exceptions > 0 && <span className="badge">{exceptions}</span>}</button>)}</nav><div className="data-status"><span>●</span> {status}</div></aside><div className="data-main"><header className="data-top"><form onSubmit={e => { e.preventDefault(); void ask(); }}><input aria-label="Ask a question" placeholder="Ask a question about your data…" value={question} onChange={e => setQuestion(e.target.value)} /><button>Ask</button></form><button className="notification-bell" type="button" aria-label={`Notifications, ${unread} unread`} aria-expanded={showNotices} onClick={() => { setShowNotices(v => !v); setUnread(0); }}><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M18 8a6 6 0 0 0-12 0c0 7-3 7-3 9h18c0-2-3-2-3-9"/><path d="M10 21h4"/></svg>{unread > 0 && <span className="unread-count">{unread > 99 ? "99+" : unread}</span>}</button></header>{understood && <div className="understood"><strong>I understood:</strong> {understood.explanation} {understood.spec && <span className="badge">{understood.pattern ?? "Query"}</span>}{!understood.ok && <details><summary>Examples</summary>{examples.map(x => <div key={x}>{x}</div>)}</details>}<button onClick={() => setUnderstood(undefined)}>×</button></div>}{error && <p role="alert">{error}</p>}<main className="data-content">{page === "Home" && <HomePage open={open} />}{page === "Dashboards" && <DashboardDesigner />}{page === "Data" && <DataPage initial={spec} onSpec={setSpec} />}{page === "Views" && <ViewsPage open={open} />}{page === "Ingestion" && <IngestionPage />}{page === "History" && <HistoryPage />}{page === "Alerts" && <AlertsPage spec={spec} />}{page === "Recipes" && <RecipesPage spec={spec} />}{page === "Sandbox" && <SandboxPage spec={spec} />}{page === "Library" && <LibraryPage />}{page === "Settings" && <SettingsPage />}</main>{showNotices && <div className="data-toasts" aria-live="polite">{notices.length ? notices.map((n, i) => <div key={i}>{n}</div>) : <div>No notifications</div>}</div>}</div></div>;
}

