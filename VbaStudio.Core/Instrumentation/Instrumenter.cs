// VbaStudio.Core/Instrumentation/Instrumenter.cs
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using VbaStudio.Core.Parsing;

namespace VbaStudio.Core.Instrumentation;

public sealed record ProbeSite(int ProbeId, string ModuleName, int OriginalLine);

public sealed record InstrumentResult(
    string InstrumentedSource,
    IReadOnlyList<ProbeSite> ProbeSites,
    int NextProbeId);

public static class Instrumenter
{
    public static InstrumentResult Instrument(
        string sourceText, ProcedureSymbols procedure, string moduleName, int startingProbeId = 1)
    {
        var physicalLines = SplitPhysicalLines(sourceText);
        var headerLine = physicalLines[procedure.StartLine - 1];
        var endLine = physicalLines[procedure.EndLine - 1];

        var bodyStartIndex = procedure.StartLine;
        var bodyEndIndexExclusive = procedure.EndLine - 1;
        var bodySlice = physicalLines.Skip(bodyStartIndex).Take(bodyEndIndexExclusive - bodyStartIndex).ToList();

        // LineJoiner already merges VBA line-continuations (" _") into one logical line, and a
        // single-line "If x Then y" is already one physical/logical line to begin with. Treating
        // each joined line as the atomic probe-placement unit therefore satisfies the instrument
        // component's "never split a continuation" and "never split a single-line If" rules for
        // free - there is never a sub-line position for a probe to land on.
        var joined = JoinPreservingComments(bodySlice);

        var probeArgs = BuildProbeArgs(procedure);
        var probeSites = new List<ProbeSite>();
        var outputLines = new List<string>();
        outputLines.AddRange(physicalLines.Take(procedure.StartLine - 1));
        outputLines.Add(headerLine);

        var currentId = startingProbeId;
        foreach (var line in joined)
        {
            var stripped = CommentStripper.StripComment(line.Text);
            var isBlank = string.IsNullOrWhiteSpace(stripped);

            if (!isBlank)
            {
                var originalLine = procedure.StartLine + line.StartPhysicalLine;
                outputLines.Add($"Agent.Probe {currentId}, Array({probeArgs})");
                probeSites.Add(new ProbeSite(currentId, moduleName, originalLine));
                currentId++;
            }

            outputLines.Add(line.Text);
        }

        outputLines.Add(endLine);
        outputLines.AddRange(physicalLines.Skip(procedure.EndLine));

        return new InstrumentResult(string.Join("\r\n", outputLines), probeSites, currentId);
    }

    // LineJoiner.Join decides whether to merge a line with its successor purely by checking
    // whether the line's trimmed text ends in " _" - it has no concept of comments. If a
    // comment's text happens to end in " _", joining on the RAW line would incorrectly merge
    // the next real statement into the comment, silently deleting it from the instrumented
    // output. This computes the join GROUPING from a comment-stripped view (so a comment's
    // trailing " _" can never trigger a false join), then reconstructs each group's text from
    // the ORIGINAL, comment-intact lines - so real comments are preserved in the instrumented
    // source exactly as written, never stripped.
    private static IReadOnlyList<JoinedLine> JoinPreservingComments(IReadOnlyList<string> rawLines)
    {
        var strippedLines = rawLines.Select(CommentStripper.StripComment).ToList();
        var groups = LineJoiner.Join(strippedLines);

        var result = new List<JoinedLine>();
        foreach (var group in groups)
        {
            var rangeLines = rawLines
                .Skip(group.StartPhysicalLine - 1)
                .Take(group.EndPhysicalLine - group.StartPhysicalLine + 1)
                .ToList();

            var sb = new StringBuilder();
            for (var i = 0; i < rangeLines.Count; i++)
            {
                if (i < rangeLines.Count - 1)
                {
                    var trimmedEnd = rangeLines[i].TrimEnd();
                    sb.Append(trimmedEnd, 0, trimmedEnd.Length - 2);
                    sb.Append(' ');
                }
                else
                {
                    sb.Append(rangeLines[i]);
                }
            }

            result.Add(new JoinedLine(sb.ToString(), group.StartPhysicalLine, group.EndPhysicalLine));
        }

        return result;
    }

    private static string BuildProbeArgs(ProcedureSymbols procedure)
    {
        var parts = new List<string>();
        foreach (var symbol in procedure.Parameters.Concat(procedure.Locals))
        {
            parts.Add($"\"{symbol.Name}\", {symbol.Name}");
        }

        return string.Join(", ", parts);
    }

    private static IReadOnlyList<string> SplitPhysicalLines(string sourceText)
    {
        var lines = Regex.Split(sourceText, "\r\n|\r|\n").ToList();
        if (lines.Count > 0 && lines[^1].Length == 0 && sourceText.Length > 0)
        {
            lines.RemoveAt(lines.Count - 1);
        }

        return lines;
    }
}
