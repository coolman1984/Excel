// Excel-style filtering and multi-key sorting over any cell source (plan 2.3 and 2.4, step 5).
// Standard library only: this file has no dependency on ExcelViewer or Excel.
using System;
using System.Collections.Generic;
using System.Threading;

namespace ExcelApp;

/// <summary>A sheet's cells as shown: text, kind (0 empty 1 number 2 text 3 bool 4 error), raw number when kind is 1.</summary>
public interface ICellSource
{
    int Rows { get; }
    int Cols { get; }
    /// <summary>Fills the arrays with one block of rows of one column. Arrays must hold at least <paramref name="count"/> items.</summary>
    void Read(int col, int fromRow, int count, string[] text, byte[] kind, double[] num);
    // Sort needs raw numbers and text, not a formatted string for every number.
    void ReadSort(int col, int fromRow, int count, string[] text, byte[] kind, double[] num) => Read(col, fromRow, count, text, kind, num);
}

/// <summary>One column's filter condition.</summary>
public abstract record Condition;
public sealed record ValuesCondition(IReadOnlySet<string> Include, bool Blanks) : Condition;
/// <summary>op: eq ne gt ge lt le between top10 aboveAverage belowAverage.</summary>
public sealed record NumberCondition(string Op, double A, double? B = null) : Condition;
/// <summary>op: eq ne contains notContains begins ends.</summary>
public sealed record TextCondition(string Op, string Value) : Condition;

/// <summary>Filters and sorts a sheet. Rows and cols here are sheet-relative 0-based, as SheetReader uses.</summary>
public sealed class FilterEngine
{
    const byte KindEmpty = 0, KindNumber = 1, KindText = 2, KindBool = 3, KindError = 4;
    const int Block = 16384;

    readonly ICellSource _source;
    readonly Dictionary<int, Condition> _conditions = new();
    IReadOnlyList<(int Col, bool Desc)> _sort = Array.Empty<(int, bool)>();

    public FilterEngine(ICellSource source) => _source = source ?? throw new ArgumentNullException(nameof(source));

    /// <summary>How many rows at the top are kept first and unfiltered. 0 when "my data has no header".</summary>
    public int HeaderRows { get; set; } = 1;

    public void SetCondition(int col, Condition? condition)
    {
        if (condition == null) _conditions.Remove(col); else _conditions[col] = condition;
    }
    public void ClearConditions() => _conditions.Clear();
    public IReadOnlyDictionary<int, Condition> Conditions => _conditions;
    public void SetSort(IReadOnlyList<(int Col, bool Desc)> keys) => _sort = keys ?? Array.Empty<(int, bool)>();

    /// <summary>The rows to show, header rows first and unfiltered. Null means the identity order (no filter, no sort).</summary>
    public int[]? Apply(CancellationToken token)
    {
        int rows = _source.Rows, cols = _source.Cols;
        int header = Math.Max(0, Math.Min(HeaderRows, rows));
        bool hasFilter = _conditions.Count > 0;
        bool hasSort = _sort.Count > 0;
        if (!hasFilter && !hasSort) return null;

        List<int> dataRows;
        if (hasFilter)
        {
            var pass = new bool[rows - header];
            Array.Fill(pass, true);
            foreach (var kv in _conditions)
            {
                token.ThrowIfCancellationRequested();
                if (kv.Key < 0 || kv.Key >= cols) continue;
                Evaluate(kv.Key, kv.Value, pass, header, rows, token);
            }
            dataRows = new List<int>(HeaderCount(pass));
            for (int r = header; r < rows; r++) if (pass[r - header]) dataRows.Add(r);
        }
        else
        {
            dataRows = new List<int>(Math.Max(0, rows - header));
            for (int r = header; r < rows; r++) dataRows.Add(r);
        }

        if (hasSort && dataRows.Count > 1) SortRows(dataRows, header, token);

        var result = new int[header + dataRows.Count];
        for (int i = 0; i < header; i++) result[i] = i;
        for (int i = 0; i < dataRows.Count; i++) result[header + i] = dataRows[i];
        return result;
    }

    static int HeaderCount(bool[] pass)
    {
        int n = 0;
        for (int i = 0; i < pass.Length; i++) if (pass[i]) n++;
        return n;
    }

    /// <summary>The distinct shown texts of one column, among the rows passing every OTHER column's
    /// condition. Sorted as Excel sorts: numbers ascending, then text A to Z, blanks counted apart.
    /// <paramref name="search"/> keeps only texts that contain it. The list is capped at <paramref name="cap"/>.</summary>
    public (List<(string Text, int Count)> Values, int Blanks, bool Truncated) Distinct(int col, string? search, int cap = 10000)
    {
        int rows = _source.Rows, cols = _source.Cols;
        int header = Math.Max(0, Math.Min(HeaderRows, rows));
        var pass = new bool[rows - header];
        Array.Fill(pass, true);
        foreach (var kv in _conditions)
            if (kv.Key != col && kv.Key >= 0 && kv.Key < cols)
                Evaluate(kv.Key, kv.Value, pass, header, rows, CancellationToken.None);

        var map = new Dictionary<string, (int Count, byte Kind, double Num)>(StringComparer.Ordinal);
        int blanks = 0;
        bool hasSearch = !string.IsNullOrEmpty(search);
        var text = new string[Block];
        var kind = new byte[Block];
        var num = new double[Block];
        int from = header;
        while (from < rows)
        {
            int count = Math.Min(Block, rows - from);
            _source.Read(col, from, count, text, kind, num);
            for (int i = 0; i < count; i++)
            {
                if (!pass[from + i - header]) continue;
                byte kd = kind[i];
                string t = text[i] ?? "";
                if (kd == KindEmpty || t.Length == 0) { blanks++; continue; }
                if (hasSearch && t.IndexOf(search!, StringComparison.OrdinalIgnoreCase) < 0) continue;
                if (map.TryGetValue(t, out var e)) map[t] = (e.Count + 1, e.Kind, e.Num);
                else map[t] = (1, kd, num[i]);
            }
            from += count;
        }

        var list = new List<(string Text, int Count)>(map.Count);
        foreach (var kv in map) list.Add((kv.Key, kv.Value.Count));
        list.Sort((x, y) =>
        {
            var a = map[x.Text];
            var b = map[y.Text];
            bool an = a.Kind == KindNumber, bn = b.Kind == KindNumber;
            if (an != bn) return an ? -1 : 1; // numbers first
            int c = an ? a.Num.CompareTo(b.Num) : string.Compare(x.Text, y.Text, StringComparison.OrdinalIgnoreCase);
            if (c == 0) c = string.CompareOrdinal(x.Text, y.Text);
            return c;
        });
        bool truncated = list.Count > cap;
        if (truncated) list = list.GetRange(0, cap);
        return (list, blanks, truncated);
    }

    void Evaluate(int col, Condition cond, bool[] pass, int header, int rows, CancellationToken token)
    {
        switch (cond)
        {
            case ValuesCondition vc:
                ForEachBlock(col, pass, header, rows, token, (k, _, t) => MatchValues(vc, k, t));
                break;
            case TextCondition tc:
                ForEachBlock(col, pass, header, rows, token, (_, __, t) => MatchText(tc, t));
                break;
            case NumberCondition nc when nc.Op is "top10" or "aboveAverage" or "belowAverage":
                EvaluateStatNumber(col, nc, pass, header, rows, token);
                break;
            case NumberCondition nc:
                ForEachBlock(col, pass, header, rows, token, (k, n, _) => MatchNumber(nc, k, n));
                break;
        }
    }

    // Runs the test on every still-passing data row, block by block, on all cores but one.
    void ForEachBlock(int col, bool[] pass, int header, int rows, CancellationToken token, Func<byte, double, string, bool> test)
    {
        int n = rows - header;
        if (n <= 0) return;
        int blocks = (n + Block - 1) / Block;
        var options = new System.Threading.Tasks.ParallelOptions
        {
            CancellationToken = token,
            MaxDegreeOfParallelism = Math.Max(1, Environment.ProcessorCount - 1)
        };
        System.Threading.Tasks.Parallel.For(0, blocks, options, b =>
        {
            int localFrom = b * Block;
            int count = Math.Min(Block, n - localFrom);
            var text = new string[count];
            var kind = new byte[count];
            var num = new double[count];
            _source.Read(col, header + localFrom, count, text, kind, num);
            for (int i = 0; i < count; i++)
            {
                int local = localFrom + i;
                if (!pass[local]) continue;
                pass[local] = test(kind[i], num[i], text[i] ?? "");
            }
        });
    }

    static bool MatchValues(ValuesCondition vc, byte kind, string text)
    {
        if (kind == KindEmpty || text.Length == 0) return vc.Blanks;
        return vc.Include != null && vc.Include.Contains(text);
    }

    static bool MatchText(TextCondition tc, string text)
    {
        string v = tc.Value ?? "";
        switch (tc.Op)
        {
            case "eq": return string.Equals(text, v, StringComparison.OrdinalIgnoreCase);
            case "ne": return !string.Equals(text, v, StringComparison.OrdinalIgnoreCase);
            case "contains": return text.IndexOf(v, StringComparison.OrdinalIgnoreCase) >= 0;
            case "notContains": return text.IndexOf(v, StringComparison.OrdinalIgnoreCase) < 0;
            case "begins": return text.StartsWith(v, StringComparison.OrdinalIgnoreCase);
            case "ends": return text.EndsWith(v, StringComparison.OrdinalIgnoreCase);
            default: return false;
        }
    }

    static bool MatchNumber(NumberCondition nc, byte kind, double num)
    {
        if (kind != KindNumber) return false; // text and empty cells fail number conditions, as in Excel
        switch (nc.Op)
        {
            case "eq": return num == nc.A;
            case "ne": return num != nc.A;
            case "gt": return num > nc.A;
            case "ge": return num >= nc.A;
            case "lt": return num < nc.A;
            case "le": return num <= nc.A;
            case "between":
                double lo = Math.Min(nc.A, nc.B ?? nc.A), hi = Math.Max(nc.A, nc.B ?? nc.A);
                return num >= lo && num <= hi;
            default: return false;
        }
    }

    // top10 / aboveAverage / belowAverage need the whole column before any row can be judged.
    void EvaluateStatNumber(int col, NumberCondition nc, bool[] pass, int header, int rows, CancellationToken token)
    {
        var nums = new List<double>();
        var text = new string[Block];
        var kind = new byte[Block];
        var num = new double[Block];
        int from = header;
        while (from < rows)
        {
            token.ThrowIfCancellationRequested();
            int count = Math.Min(Block, rows - from);
            _source.Read(col, from, count, text, kind, num);
            for (int i = 0; i < count; i++) if (kind[i] == KindNumber) nums.Add(num[i]);
            from += count;
        }
        if (nc.Op == "aboveAverage" || nc.Op == "belowAverage")
        {
            double sum = 0;
            foreach (var d in nums) sum += d;
            double avg = nums.Count > 0 ? sum / nums.Count : 0;
            bool above = nc.Op == "aboveAverage";
            ForEachBlock(col, pass, header, rows, token, (k, n, _) => k == KindNumber && (above ? n > avg : n < avg));
        }
        else // top10: the top A items (10 by default), ties at the cut included
        {
            nums.Sort();
            int take = nc.A > 0 ? (int)nc.A : 10;
            double threshold = nums.Count > 0 ? nums[Math.Max(0, nums.Count - take)] : double.PositiveInfinity;
            ForEachBlock(col, pass, header, rows, token, (k, n, _) => k == KindNumber && n >= threshold);
        }
    }

    // Stable, multi-key sort: the first key decides, later keys break its ties, original order breaks the last.
    void SortRows(List<int> rows, int header, CancellationToken token)
    {
        int n = rows.Count, keys = _sort.Count, cols = _source.Cols;
        int span = _source.Rows - header;
        var kinds = new byte[keys][];
        var nums = new double[keys][];
        var texts = new string[keys][];
        var bt = new string[Block];
        var bk = new byte[Block];
        var bn = new double[Block];
        for (int k = 0; k < keys; k++)
        {
            var kind = new byte[span];
            var num = new double[span];
            var text = new string[span];
            int col = _sort[k].Col;
            if (col >= 0 && col < cols)
            {
                int from = header;
                while (from < _source.Rows)
                {
                    token.ThrowIfCancellationRequested();
                    int count = Math.Min(Block, _source.Rows - from);
                    _source.ReadSort(col, from, count, bt, bk, bn);
                    for (int i = 0; i < count; i++)
                    {
                        int local = from + i - header;
                        byte kd = bk[i];
                        string tx = bt[i] ?? "";
                        if (kd == KindText && tx.Length == 0) kd = KindEmpty;
                        kind[local] = kd;
                        num[local] = bn[i];
                        text[local] = tx;
                    }
                    from += count;
                }
            }
            kinds[k] = kind;
            nums[k] = num;
            texts[k] = text;
        }

        // Array.Sort's numeric key/value path avoids a managed callback for every comparison.
        // Restore row order within equal-value runs to keep Excel's stable tie ordering.
        if (keys == 1 && rows.TrueForAll(r => kinds[0][r - header] is KindNumber or KindEmpty))
        {
            var numericRows = rows.FindAll(r => kinds[0][r - header] == KindNumber).ToArray();
            var values = new double[numericRows.Length];
            for (int i = 0; i < values.Length; i++) values[i] = nums[0][numericRows[i] - header];
            Array.Sort(values, numericRows);
            if (_sort[0].Desc) { Array.Reverse(values); Array.Reverse(numericRows); }
            for (int start = 0; start < values.Length;) {
                token.ThrowIfCancellationRequested();
                int end = start + 1;
                while (end < values.Length && values[end].CompareTo(values[start]) == 0) end++;
                if (end - start > 1) Array.Sort(numericRows, start, end - start);
                start = end;
            }
            var emptyRows = rows.FindAll(r => kinds[0][r - header] == KindEmpty);
            rows.Clear(); rows.AddRange(numericRows); rows.AddRange(emptyRows);
            return;
        }

        var order = new int[n];
        for (int i = 0; i < n; i++) order[i] = i;
        var sort = _sort;
        Array.Sort(order, (pa, pb) =>
        {
            int ra = rows[pa] - header, rb = rows[pb] - header;
            for (int k = 0; k < keys; k++)
            {
                int c = CompareKey(kinds[k][ra], nums[k][ra], texts[k][ra],
                                   kinds[k][rb], nums[k][rb], texts[k][rb], sort[k].Desc);
                if (c != 0) return c;
            }
            return pa.CompareTo(pb); // stability
        });
        var reordered = new List<int>(n);
        for (int i = 0; i < n; i++) reordered.Add(rows[order[i]]);
        for (int i = 0; i < n; i++) rows[i] = reordered[i];
    }

    // Same order as ExcelViewer SortEngine: numbers, then text, then booleans, then errors; empties last.
    // Descending reverses everything but the empties.
    static int CompareKey(byte ka, double na, string ta, byte kb, double nb, string tb, bool desc)
    {
        bool ea = ka == KindEmpty, eb = kb == KindEmpty;
        if (ea || eb)
        {
            if (ea && eb) return 0;
            return ea ? 1 : -1;
        }
        int c = ka.CompareTo(kb);
        if (c == 0)
            c = ka == KindText ? string.Compare(ta, tb, StringComparison.OrdinalIgnoreCase) : na.CompareTo(nb);
        if (desc) c = -c;
        return c;
    }
}
