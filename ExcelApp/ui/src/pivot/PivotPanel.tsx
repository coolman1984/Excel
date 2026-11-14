// The PivotTable panel: pick fields for Rows, Columns and Values (sum, count, average, min, max); the summary updates as
// you choose. "Place on sheet" writes it next to the data as values (one undo step, saved like any edit).
import * as React from "react";
import { call } from "../bridge";
import { fullNumber } from "../charts/scale";

interface Field { name: string; type: "number" | "text" }
interface Value { field: string; fn: string }
interface Result { rowFields: string[]; columnField: string | null; valueLabels: string[]; columnKeys: string[]; rows: { keys: string[]; cells: (number | null)[][]; total: (number | null)[] }[]; columnTotals: (number | null)[][]; grand: (number | null)[]; truncated: boolean }
const FNS: [string, string][] = [["sum", "Sum"], ["count", "Count"], ["avg", "Average"], ["min", "Min"], ["max", "Max"]];

export function PivotPanel({ sheet, range, onClose, onPlaced }: { sheet: number; range: number[]; onClose: () => void; onPlaced: (where: string) => void }): React.ReactElement {
    const [fields, setFields] = React.useState<Field[]>([]);
    const [source, setSource] = React.useState("");
    const [rows, setRows] = React.useState<string[]>([]);
    const [column, setColumn] = React.useState<string | null>(null);
    const [values, setValues] = React.useState<Value[]>([]);
    const [result, setResult] = React.useState<Result | null>(null);
    const [error, setError] = React.useState<string | null>(null);
    const [busy, setBusy] = React.useState(false);

    React.useEffect(() => {
        call<any>("pivot.fields", { sheet, range }).then((f) => {
            setFields(f.fields ?? []); setSource(f.range ?? "");
            // a sensible start, as Excel suggests: the first text field in Rows, the first number field summed
            const text = (f.fields ?? []).find((x: Field) => x.type === "text"), num = (f.fields ?? []).find((x: Field) => x.type === "number");
            if (text) setRows([text.name]);
            if (num) setValues([{ field: num.name, fn: "sum" }]);
        }).catch((e) => setError(e?.message ?? String(e)));
    }, [sheet, range]);

    const spec = { sheet, range, rows, column, values };
    React.useEffect(() => {
        if ((rows.length === 0 && !column) || values.length === 0) { setResult(null); return; }
        let live = true; setBusy(true);
        const t = setTimeout(() => {
            call<Result>("pivot.run", spec).then((r) => { if (live) { setResult(r); setError(null); } })
                .catch((e) => { if (live) { setResult(null); setError(e?.message ?? String(e)); } }).finally(() => live && setBusy(false));
        }, 150);
        return () => { live = false; clearTimeout(t); };
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [JSON.stringify(spec)]);

    const toggleRow = (f: string) => { setRows((r) => (r.includes(f) ? r.filter((x) => x !== f) : r.length < 3 ? [...r, f] : r)); if (column === f) setColumn(null); };
    const toggleCol = (f: string) => { setColumn((c) => (c === f ? null : f)); setRows((r) => r.filter((x) => x !== f)); };
    const addValue = (f: Field) => setValues((v) => [...v, { field: f.name, fn: f.type === "number" ? "sum" : "count" }]);
    const place = async () => {
        try { const r: any = await call("pivot.place", spec); onPlaced(r.placed); }
        catch (e: any) { setError(e?.message ?? String(e)); }
    };

    return (
        <aside className="pivotPanel" aria-label="PivotTable" data-testid="pivot-panel">
            <header className="agentHead"><strong>PivotTable</strong><span className="pivotSource">{source}</span>
                <button type="button" className="agentClose" onClick={onClose} aria-label="Close the PivotTable panel">×</button></header>
            <div className="pivotBody">
                <section className="pivotFields" aria-label="Fields">
                    <h3>Fields</h3>
                    {fields.map((f) => (
                        <div key={f.name} className="pivotField">
                            <span className="pivotFieldName" title={f.type === "number" ? "Numbers" : "Text"}>{f.type === "number" ? "#" : "Aa"} {f.name}</span>
                            <button type="button" className={rows.includes(f.name) ? "on" : ""} onClick={() => toggleRow(f.name)} aria-pressed={rows.includes(f.name)} title="Rows">Rows</button>
                            <button type="button" className={column === f.name ? "on" : ""} onClick={() => toggleCol(f.name)} aria-pressed={column === f.name} title="Columns">Cols</button>
                            <button type="button" onClick={() => addValue(f)} title="Add to Values">Σ</button>
                        </div>
                    ))}
                    {values.length > 0 && <h3>Values</h3>}
                    {values.map((v, i) => (
                        <div key={i} className="pivotValue">
                            <select aria-label={`Summarise ${v.field} by`} value={v.fn} onChange={(e) => setValues((all) => all.map((x, j) => (j === i ? { ...x, fn: e.target.value } : x)))}>
                                {FNS.filter(([k]) => k === "count" || fields.find((f) => f.name === v.field)?.type === "number").map(([k, l]) => <option key={k} value={k}>{l}</option>)}
                            </select>
                            <span>of {v.field}</span>
                            <button type="button" aria-label={`Remove ${v.field}`} onClick={() => setValues((all) => all.filter((_, j) => j !== i))}>×</button>
                        </div>
                    ))}
                </section>
                <section className="pivotResult" aria-label="Result">
                    {error && <p className="pivotError" role="alert">{error}</p>}
                    {!error && !result && <p className="agentHint">{busy ? "Calculating..." : "Choose a field for Rows (or Columns) and one for Values."}</p>}
                    {result && <PivotTableView r={result} />}
                </section>
            </div>
            <footer className="pivotFoot">
                {result?.truncated && <span className="pivotNote">Showing the first groups only.</span>}
                <button type="button" className="agentSend" disabled={!result} onClick={() => void place()}>Place on sheet</button>
            </footer>
        </aside>
    );
}

function PivotTableView({ r }: { r: Result }): React.ReactElement {
    const nv = r.valueLabels.length, cols = r.columnField != null;
    const rowHead = r.rowFields.length ? r.rowFields : [""];
    return (
        <div className="pivotTableWrap">
            <table className="pivotTable">
                <thead>
                    {cols && <tr>{rowHead.map((_, i) => <th key={i} className="pivotKey">{i === 0 && nv === 1 ? r.valueLabels[0] : ""}</th>)}{r.columnKeys.map((k) => <th key={k} colSpan={nv} className="pivotColKey">{k}</th>)}<th colSpan={nv}>Grand Total</th></tr>}
                    <tr>{rowHead.map((f, i) => <th key={i}>{f}</th>)}
                        {cols ? [...r.columnKeys, "total"].flatMap((k) => r.valueLabels.map((v) => <th key={k + v}>{nv === 1 ? (cols ? "" : v) : v}</th>)) : r.valueLabels.map((v) => <th key={v}>{v}</th>)}</tr>
                </thead>
                <tbody>
                    {r.rows.slice(0, 500).map((row, i) => (
                        <tr key={i}>{row.keys.map((k, j) => <td key={j} className="pivotKey">{k}</td>)}
                            {cols ? [...row.cells, row.total].flatMap((c, j) => c.map((v, k) => <td key={j + ":" + k}>{fullNumber(v)}</td>)) : row.total.map((v, k) => <td key={k}>{fullNumber(v)}</td>)}</tr>
                    ))}
                    <tr className="pivotGrand"><td className="pivotKey">Grand Total</td>{rowHead.slice(1).map((_, i) => <td key={"e" + i} />)}
                        {cols ? [...r.columnTotals, r.grand].flatMap((c, j) => c.map((v, k) => <td key={j + ":" + k}>{fullNumber(v)}</td>)) : r.grand.map((v, k) => <td key={k}>{fullNumber(v)}</td>)}</tr>
                </tbody>
            </table>
        </div>
    );
}
