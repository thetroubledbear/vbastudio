namespace VbaStudio.Core.Tooling;

public static class StaleChecker
{
    // A module is "stale" (from a run-safety perspective) if its on-disk copy doesn't exist at
    // all (never pulled) or differs from what's currently in Excel (edited but not pushed, or
    // Excel changed since the last pull) - both cases mean "what's about to run may not be what
    // you're looking at on disk," which is the one thing this check exists to catch.
    public static bool IsStale(string? diskContent, string excelContent)
    {
        return diskContent == null || diskContent != excelContent;
    }
}
