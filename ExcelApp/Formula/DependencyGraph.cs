namespace ExcelApp.Formula;

public readonly record struct Cell(int Sheet, int Row, int Col);
public readonly record struct CellRange(int Sheet, int FirstRow, int LastRow, int FirstCol, int LastCol)
{
    public bool Contains(Cell cell) => cell.Sheet == Sheet && cell.Row >= FirstRow && cell.Row <= LastRow && cell.Col >= FirstCol && cell.Col <= LastCol;
}
public sealed record Recalculation(IReadOnlyList<Cell> Ordered, IReadOnlySet<Cell> Circular);
public sealed class DependencyGraph
{
    readonly Dictionary<Cell, CellRange[]> formulas = new();
    sealed class RangeIndex
    {
        readonly List<(Cell cell, CellRange range)> entries = new(); int[] prefixMax = Array.Empty<int>(); bool dirty;
        public int Count => entries.Count;
        public void Add((Cell cell, CellRange range) item) { entries.Add(item); dirty = true; }
        public void Remove(Cell cell, CellRange range) { entries.RemoveAll(x => x.cell == cell && x.range == range); dirty = true; }
        public IEnumerable<Cell> Find(int row)
        {
            if (dirty) { entries.Sort((a, b) => a.range.FirstRow.CompareTo(b.range.FirstRow)); prefixMax = new int[entries.Count]; int maximum = -1; for (int i = 0; i < entries.Count; i++) { maximum = Math.Max(maximum, entries[i].range.LastRow); prefixMax[i] = maximum; } dirty = false; }
            int lo = 0, hi = entries.Count; while (lo < hi) { int mid = lo + (hi - lo) / 2; if (entries[mid].range.FirstRow <= row) lo = mid + 1; else hi = mid; }
            for (int i = lo - 1; i >= 0 && prefixMax[i] >= row; i--) if (entries[i].range.LastRow >= row) yield return entries[i].cell;
        }
    }
    readonly Dictionary<(int sheet, int col), RangeIndex> index = new();
    public void Set(Cell cell, IEnumerable<CellRange> dependencies)
    {
        Remove(cell); var ranges = dependencies.ToArray(); formulas[cell] = ranges;
        foreach (var range in ranges) for (int col = range.FirstCol; col <= range.LastCol; col++) { var key = (range.Sheet, col); if (!index.TryGetValue(key, out var list)) index[key] = list = new RangeIndex(); list.Add((cell, range)); }
    }
    public void Remove(Cell cell)
    {
        if (!formulas.Remove(cell, out var ranges)) return;
        foreach (var range in ranges) for (int col = range.FirstCol; col <= range.LastCol; col++) { var key = (range.Sheet, col); if (index.TryGetValue(key, out var list)) { list.Remove(cell, range); if (list.Count == 0) index.Remove(key); } }
    }
    IEnumerable<Cell> Dependents(Cell cell) => index.TryGetValue((cell.Sheet, cell.Col), out var list) ? list.Find(cell.Row).Distinct() : Enumerable.Empty<Cell>();
    public Recalculation Recalculate(IEnumerable<Cell> changes)
    {
        var affected = new HashSet<Cell>(); var queue = new Queue<Cell>(changes); var expanded = new HashSet<Cell>();
        while (queue.TryDequeue(out var source)) { if (!expanded.Add(source)) continue; foreach (var dest in Dependents(source)) if (affected.Add(dest)) queue.Enqueue(dest); }
        var indegree = affected.ToDictionary(x => x, _ => 0); var edges = new Dictionary<Cell, List<Cell>>();
        foreach (var source in affected) foreach (var dest in Dependents(source)) if (affected.Contains(dest)) { if (!edges.TryGetValue(source, out var list)) edges[source] = list = new(); list.Add(dest); indegree[dest]++; }
        var ready = new Queue<Cell>(indegree.Where(x => x.Value == 0).Select(x => x.Key)); var result = new List<Cell>(affected.Count);
        while (ready.TryDequeue(out var node)) { result.Add(node); if (edges.TryGetValue(node, out var targets)) foreach (var target in targets) if (--indegree[target] == 0) ready.Enqueue(target); }
        var remaining = indegree.Where(x => x.Value > 0).Select(x => x.Key).ToHashSet();
        var finish = new List<Cell>(); var visited = new HashSet<Cell>();
        foreach (var start in remaining)
        {
            if (!visited.Add(start)) continue; var stack = new Stack<(Cell cell, int next)>(); stack.Push((start, 0));
            while (stack.TryPop(out var frame))
            {
                if (!edges.TryGetValue(frame.cell, out var targets) || frame.next >= targets.Count) { finish.Add(frame.cell); continue; }
                var target = targets[frame.next]; stack.Push((frame.cell, frame.next + 1)); if (remaining.Contains(target) && visited.Add(target)) stack.Push((target, 0));
            }
        }
        var reverse = new Dictionary<Cell, List<Cell>>(); foreach (var entry in edges) if (remaining.Contains(entry.Key)) foreach (var target in entry.Value) if (remaining.Contains(target)) { if (!reverse.TryGetValue(target, out var list)) reverse[target] = list = new(); list.Add(entry.Key); }
        visited.Clear(); var circular = new HashSet<Cell>();
        for (int i = finish.Count - 1; i >= 0; i--)
        {
            var start = finish[i]; if (!visited.Add(start)) continue; var component = new List<Cell>(); var stack = new Stack<Cell>(); stack.Push(start);
            while (stack.TryPop(out var node)) { component.Add(node); if (reverse.TryGetValue(node, out var parents)) foreach (var parent in parents) if (visited.Add(parent)) stack.Push(parent); }
            if (component.Count > 1 || (edges.TryGetValue(start, out var targets) && targets.Contains(start))) foreach (var node in component) circular.Add(node);
        }
        var residualDegree = remaining.Where(x => !circular.Contains(x)).ToDictionary(x => x, _ => 0);
        foreach (var entry in edges) if (residualDegree.ContainsKey(entry.Key)) foreach (var target in entry.Value) if (residualDegree.ContainsKey(target)) residualDegree[target]++;
        var residualReady = new Queue<Cell>(residualDegree.Where(x => x.Value == 0).Select(x => x.Key));
        while (residualReady.TryDequeue(out var node)) { result.Add(node); if (edges.TryGetValue(node, out var targets)) foreach (var target in targets) if (residualDegree.ContainsKey(target) && --residualDegree[target] == 0) residualReady.Enqueue(target); }
        return new Recalculation(result, circular);
    }
}
