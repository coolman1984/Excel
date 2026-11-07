using System.Globalization;

namespace ExcelApp;

/** A capped probe is an incomplete observation, never a five-second latency. */
public static class JumpSamples
{
    public static string Format(double? milliseconds) => milliseconds is double ms
        ? ms.ToString("F0", CultureInfo.InvariantCulture) + " ms" : "TIMEOUT";

    public static string Summary(IReadOnlyList<double?> samples)
    {
        if (samples.Count == 0) return "JUMP no samples";
        // Incomplete observations sort after completed ones, keeping TIMEOUT in
        // the relevant percentile/worst rather than disguising it as 5000 ms.
        var sorted = samples.OrderBy(x => x ?? double.PositiveInfinity).ToArray();
        int timeouts = samples.Count(x => x == null);
        return $"JUMP median {Format(sorted[sorted.Length / 2])}, p95 {Format(sorted[(int)(sorted.Length * 0.95)])}, worst {Format(sorted[^1])}; TIMEOUT {timeouts}/{samples.Count}";
    }
}
