// VbaStudio.Core/Instrumentation/Instrumenter.cs
using System.Collections.Generic;
using System.Linq;
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
        var joined = LineJoiner.Join(bodySlice);

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
