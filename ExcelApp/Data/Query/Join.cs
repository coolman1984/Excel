using System.Globalization;
using System.Text;

namespace ExcelApp.Data.Query;

public static class Join
{
    public static Table Run(Table left, Table right, JoinSpec spec, CancellationToken token)
    {
        if (spec.On.Count == 0) throw new ArgumentException("A join needs at least one key.");
        if (spec.Kind is not ("inner" or "left" or "anti")) throw new ArgumentException("Unknown join kind.");
        string mode = spec.Fuzzy ?? "none"; if (mode is not ("none" or "normalized" or "similar")) throw new ArgumentException("Unknown matching mode.");
        var lc = spec.On.Select(k => left[k.Left] ?? throw new ArgumentException("Unknown join column.")).ToArray();
        var rc = spec.On.Select(k => right[k.Right] ?? throw new ArgumentException("Unknown join column.")).ToArray();
        string? Key(Column[] columns, int row, bool normalized)
        {
            var builder = new StringBuilder();
            foreach (var c in columns) { if (c.IsNull(row)) return null; string value = c.Text(row); if (normalized) value = Fuzzy.Normalize(value); else value = value.ToUpperInvariant(); builder.Append(value.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(value); }
            return builder.ToString();
        }
        string SimilarKey(Column[] columns, int row) => string.Join("\u001f", columns.Select(c => Fuzzy.Normalize(c.Text(row))));
        bool norm = mode != "none";
        var index = new Dictionary<string, List<int>>(StringComparer.Ordinal); Dictionary<(char, int), List<string>>? buckets = mode == "similar" ? new() : null;
        for (int r = 0; r < right.Rows; r++)
        {
            if ((r & 65535) == 0) token.ThrowIfCancellationRequested(); var key = Key(rc, r, norm); if (key == null) continue;
            if (!index.TryGetValue(key, out var list)) { index[key] = list = new List<int>(1); if (buckets != null) { var value = SimilarKey(rc, r); if (value.Length > 0) { var bucketKey = (value[0], value.Length); if (!buckets.TryGetValue(bucketKey, out var bucket)) buckets[bucketKey] = bucket = new(); bucket.Add(key); } } } list.Add(r);
        }
        var leftRows = new List<int>(); var rightRows = new List<int>(); var matches = new List<string?>(); var scores = new List<double>();
        for (int l = 0; l < left.Rows; l++)
        {
            if ((l & 65535) == 0) token.ThrowIfCancellationRequested(); string? key = Key(lc, l, norm); List<int>? found = null; double score = 1; bool similar = false;
            if (key != null) index.TryGetValue(key, out found);
            if (found == null && mode == "similar" && key != null )
            {
                string value = SimilarKey(lc, l); double best = spec.MinScore ?? .92; string? bestKey = null;
                if (value.Length > 0) for (int n = Math.Max(1, value.Length - 2); n <= value.Length + 2; n++) if (buckets!.TryGetValue((value[0], n), out var candidates)) foreach (string candidate in candidates) { string other = SimilarKey(rc, index[candidate][0]); double current = Fuzzy.JaroWinkler(value, other); if (current >= best && (bestKey == null || current > best)) { best = current; bestKey = candidate; } }
                if (bestKey != null) { found = index[bestKey]; score = best; similar = true; }
            }
            if (found == null) { if (spec.Kind != "inner") { leftRows.Add(l); rightRows.Add(-1); matches.Add("none"); scores.Add(0); } continue; }
            if (spec.Kind == "anti") continue;
            string? raw = Key(lc, l, false);
            foreach (int r in found) { leftRows.Add(l); rightRows.Add(r); matches.Add(similar ? "similar" : raw == Key(rc, r, false) ? "exact" : "normalized"); scores.Add(score); }
        }
        var leftIndex = leftRows.ToArray(); var rightIndex = rightRows.ToArray(); var output = left.Columns.Select(c => Engine.Gather(c, leftIndex)).ToList();
        if (spec.Kind != "anti") foreach (var col in right.Columns)
        {
            string name = (spec.Prefix ?? "") + col.Name; if (output.Any(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase))) name = "right." + name;
            output.Add(GatherRight(col, name, rightIndex));
        }
        output.Add(TextColumn.From("_match", matches)); output.Add(new NumberColumn("_score", scores.ToArray(), null)); return new Table(left.Name, output);
    }
    static Column GatherRight(Column col, string name, int[] rows)
    {
        if (col is TextColumn text) return new TextColumn(name, rows.Select(r => r < 0 ? -1 : text.Codes[r]).ToArray(), text.Dict);
        if (col is BoolColumn boolean) return new BoolColumn(name, rows.Select(r => r < 0 ? (sbyte)-1 : boolean.Values[r]).ToArray());
        return new NumberColumn(name, rows.Select(r => r < 0 ? 0 : col.Num(r)).ToArray(), rows.Select(r => r < 0 || col.IsNull(r)).ToArray(), col.Type);
    }
}
