namespace VbaStudio.Core.Tooling;

public static class StaleChecker
{
    // A module is "stale" (from a run-safety perspective) if its on-disk copy doesn't exist at
    // all (never pulled) or differs from what's currently in Excel (edited but not pushed, or
    // Excel changed since the last pull) - both cases mean "what's about to run may not be what
    // you're looking at on disk," which is the one thing this check exists to catch.
    //
    // The comparison is normalized (line endings unified to \n, trailing whitespace trimmed)
    // rather than byte-exact. A byte-exact compare turns any trailing-newline or line-ending
    // difference - which editors add routinely (insertFinalNewline, formatters, git autocrlf) -
    // into a permanent false "stale", one the user cannot clear: "Push and run" pushes the disk
    // text into Excel, but Excel's own export re-strips the trailing newline every time, so the
    // next check reports stale again. Whitespace-only differences at the edges of the file are
    // never a run-safety concern, so they are normalized away instead.
    public static bool IsStale(string? diskContent, string excelContent)
    {
        return diskContent == null || Normalize(diskContent) != Normalize(excelContent);
    }

    public static string Normalize(string content)
        => content.Replace("\r\n", "\n").Replace("\r", "\n").TrimEnd();
}
