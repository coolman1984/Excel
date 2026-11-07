// Search and sort. Both run in the background on the copy, with no Excel, and never change the source.
using System.Text;

namespace ExcelViewer.Core;

public readonly record struct Hit(int Sheet, int Row, int Col);

public static class SearchEngine
{
    public const int MaxHits = 10000;

    /// <summary>Finds text in text cells, and in numbers as they are shown. Uses all cores but one; hits arrive in order.
    /// Returns true if it covered everything.</summary>
    public static bool Run(CacheData cache, IList<int> sheets, string query, Action<Hit> onHit, Action<int> onPercent, CancellationToken token)
    {
        if (string.IsNullOrEmpty(query)) return true;
        long total = Math.Max(1, sheets.Sum(s => cache.Sheets[s].Cells)), seen = 0; int found = 0, lastPct = -1; bool capped = false;
        bool digits = query.Any(char.IsDigit); string plainQuery = query.Replace(",", "");
        var options = new ParallelOptions { CancellationToken = token, MaxDegreeOfParallelism = Math.Max(1, Environment.ProcessorCount - 1) };
        foreach (int s in sheets)
        {
            if (capped) break;
            var d = cache.Sheets[s]; int chunks = d.Chunks.Count; var done = new List<Hit>[chunks]; int next = 0; var gate = new object();
            using var readers = new ThreadLocal<SheetReader>(() => new SheetReader(cache, d), true);
            try
            {
                Parallel.For(0, chunks, options, (ci, state) =>
                {
                    if (token.IsCancellationRequested || Volatile.Read(ref capped)) { state.Stop(); return; }
                    var list = Scan(readers.Value, s, ci, query, plainQuery, digits);
                    lock (gate)
                    {
                        done[ci] = list; seen += (long)readers.Value.RowsInChunk(ci) * d.Cols;
                        while (next < chunks && done[next] != null && !capped)
                        {
                            foreach (var h in done[next]) { onHit(h); if (++found >= MaxHits) { capped = true; state.Stop(); break; } }
                            done[next] = null; next++; if (next < chunks && done[next] == null) break;
                        }
                        int pct = (int)(seen * 100 / total); if (pct != lastPct) { lastPct = pct; onPercent(pct); }
                    }
                });
            }
            finally { foreach (var r in readers.Values) r.Dispose(); }
        }
        return !capped;
    }

    static List<Hit> Scan(SheetReader rd, int sheet, int ci, string query, string plainQuery, bool digits)
    {
        var d = rd.Sheet; var hits = new List<Hit>(0); int n = rd.RowsInChunk(ci), r0 = ci * d.BlockRows;
        for (int k = 0; k < d.Cols; k++)
        {
            var c = rd.LoadColumn(ci, k);
            bool[] entry = null;
            if (c.Blob != null && c.Index != null)
            {
                entry = new bool[c.Offs.Length - 1];
                for (int e = 0; e < entry.Length; e++) entry[e] = Encoding.UTF8.GetString(c.Blob, c.Offs[e], c.Offs[e + 1] - c.Offs[e]).Contains(query, StringComparison.OrdinalIgnoreCase);
            }
            bool textMayHit = c.Blob != null && (entry == null || entry.Any(x => x));
            if (!textMayHit && !(digits && c.Nums != null)) continue;
            for (int i = 0; i < n; i++)
            {
                byte kd = c.KindAt(i); bool hit = false;
                if (kd == CellKind.Text && textMayHit) hit = entry != null ? entry[c.Index[i]] : c.TextAt(i).Contains(query, StringComparison.OrdinalIgnoreCase);
                else if (kd == CellKind.Number && digits)
                {
                    string t = rd.Show(c, i, r0 + i, k).Text;
                    hit = t.Contains(query, StringComparison.OrdinalIgnoreCase) || t.Replace(",", "").Contains(plainQuery, StringComparison.OrdinalIgnoreCase);
                }
                if (hit) hits.Add(new Hit(sheet, r0 + i, k));
            }
        }
        return hits;
    }
}

public static class SortEngine
{
    /// <summary>A new order of rows. Numbers, then text, then TRUE/FALSE, then errors, with empty cells always last.</summary>
    public static int[] Sort(CacheData cache, SheetInfo d, int col, bool descending, int firstDataRow, CancellationToken token)
    {
        return Sort(cache, d, col, descending, Enumerable.Range(firstDataRow, Math.Max(0, d.Rows - firstDataRow)).ToArray(), token);
    }

    /// <summary>Stable in the supplied row order, including ties and empty cells.</summary>
    public static int[] Sort(CacheData cache, SheetInfo d, int col, bool descending, IReadOnlyList<int> rows, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (col < 0 || col >= d.Cols) throw new ArgumentOutOfRangeException(nameof(col));
        foreach (int row in rows) if (row < 0 || row >= d.Rows) throw new ArgumentOutOfRangeException(nameof(rows));
        if (rows.Count == 0) return Array.Empty<int>();
        var values = new (int row, int order, byte kind, double num, string text)[rows.Count];
        using var rd = new SheetReader(cache, d);
        // Read each chunk once; the order field preserves an earlier sort's ties.
        foreach (var group in rows.Select((row, order) => (row, order)).GroupBy(x => x.row / d.BlockRows))
        {
            token.ThrowIfCancellationRequested();
            var chunk = rd.LoadColumn(group.Key, col);
            foreach (var item in group)
            {
                if (item.row < 0 || item.row >= d.Rows) throw new ArgumentOutOfRangeException(nameof(rows));
                int i = item.row % d.BlockRows; byte kind = chunk.KindAt(i);
                string text = kind == CellKind.Text ? chunk.TextAt(i) : null;
                if (text == "") kind = CellKind.Empty;
                values[item.order] = (item.row, item.order, kind, chunk.Nums == null ? 0 : chunk.Nums[i], text);
            }
        }
        Array.Sort(values, (a, b) => {
            int cmp;
            if (a.kind == CellKind.Empty || b.kind == CellKind.Empty)
                cmp = a.kind == b.kind ? 0 : a.kind == CellKind.Empty ? 1 : -1;
            else {
                cmp = a.kind.CompareTo(b.kind);
                if (cmp == 0) cmp = a.kind == CellKind.Text ? string.Compare(a.text, b.text, StringComparison.OrdinalIgnoreCase) : a.num.CompareTo(b.num);
                if (descending) cmp = -cmp;
            }
            return cmp == 0 ? a.order.CompareTo(b.order) : cmp;
        });
        token.ThrowIfCancellationRequested();
        return values.Select(x => x.row).ToArray();
    }
}

/// <summary>Builds the text for the clipboard: tabs between columns, line breaks between rows, the way Excel pastes it.</summary>
public static class CopyText
{
    public const int MaxCells = 200_000;

    public static string Build(IEnumerable<(int Row, int Col)> cells, Func<int, int, string> shown)
    {
        var chosen = cells.ToHashSet(); if (chosen.Count == 0) return "";
        var rows = chosen.Select(c => c.Row).Distinct().OrderBy(x => x).ToList();
        var cols = chosen.Select(c => c.Col).Distinct().OrderBy(x => x).ToList();
        var sb = new StringBuilder();
        for (int r = 0; r < rows.Count; r++)
        {
            if (r > 0) sb.Append("\r\n");
            for (int c = 0; c < cols.Count; c++)
            {
                if (c > 0) sb.Append('\t');
                if (chosen.Contains((rows[r], cols[c]))) sb.Append(Quote(shown(rows[r], cols[c])));
            }
        }
        return sb.ToString();
    }

    // A tab or line break inside a cell would split it when pasted, so such a cell is quoted, as Excel does.
    static string Quote(string s) => s.IndexOfAny(new[] { '\t', '\n', '\r' }) < 0 ? s : "\"" + s.Replace("\"", "\"\"") + "\"";
}
