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
    // A "With"/"End With" line is itself always skipped - probing right as we enter or leave
    // the block reveals nothing a neighboring probe doesn't already show. Every regular line
    // while withDepth > 0 is skipped too, since the plan explicitly prefers skipping a With
    // block's body over attempting to capture its target expression (real expression
    // evaluation this component doesn't have).
    private static readonly Regex WithPattern = new(@"^\s*With\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex EndWithPattern = new(@"^\s*End\s+With\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // VBA has exactly one position where inserting any statement is a hard compile error:
    // between "Select Case <expr>" and its first "Case" clause. The Select Case line itself
    // still gets probed normally (nothing wrong with probing before it runs) - only the next
    // probeable line, which valid VBA syntax guarantees is the first Case clause, is suppressed.
    private static readonly Regex SelectCasePattern = new(@"^\s*Select\s+Case\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

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

        // Parameters are valid from procedure entry, but a local is only valid to reference
        // from its own Dim/Static/Const line onward - Option Explicit makes VBA reject a
        // forward reference to a not-yet-declared local as a compile error, confirmed live
        // against real Excel. inScope grows as declaration lines are passed, so each probe
        // only ever names variables that are actually in scope at that point.
        var inScope = new List<Symbol>(procedure.Parameters);
        var probeSites = new List<ProbeSite>();
        var outputLines = new List<string>();
        outputLines.AddRange(physicalLines.Take(procedure.StartLine - 1));
        outputLines.Add(headerLine);

        var currentId = startingProbeId;
        var withDepth = 0;
        var suppressNextProbe = false;
        var previousEmittedEndsInContinuationMarker = false;
        foreach (var line in joined)
        {
            var stripped = CommentStripper.StripComment(line.Text);
            var isWithLine = WithPattern.IsMatch(stripped);
            var isEndWithLine = EndWithPattern.IsMatch(stripped);
            var isSelectCaseLine = SelectCasePattern.IsMatch(stripped);

            bool skip;
            if (isWithLine)
            {
                skip = true;
                withDepth++;
            }
            else if (isEndWithLine)
            {
                skip = true;
                if (withDepth > 0)
                {
                    withDepth--;
                }
            }
            else
            {
                skip = withDepth > 0;
            }

            var isBlank = string.IsNullOrWhiteSpace(stripped);

            // We cannot empirically confirm (no Excel in this milestone) whether real VBA
            // treats a trailing " _" inside a comment as continuing that comment onto the next
            // physical line - there is credible evidence it does. If it does, a probe line
            // emitted immediately after such a line risks being silently swallowed into the
            // continued comment, so it would never fire. Suppressing that probe is safe under
            // either reading of the open question.
            if (previousEmittedEndsInContinuationMarker && !skip && !isBlank)
            {
                skip = true;
            }

            // VBA forbids any statement between "Select Case <expr>" and its first "Case"
            // clause - inserting a probe there is a hard compile error, not just semantically
            // odd. The Select Case line itself still gets probed normally (nothing wrong with
            // probing before it runs); only the next probeable line - which VBA syntax
            // guarantees is the first Case clause - is suppressed.
            if (suppressNextProbe && !isBlank)
            {
                skip = true;
                suppressNextProbe = false;
            }

            if (!skip && !isBlank)
            {
                var originalLine = procedure.StartLine + line.StartPhysicalLine;
                // "Agent.Probe" refers to vba/modAgent.bas, built as part of this same milestone.
                // It exposes Public Sub Probe(id As Long, values As Variant) - a plain Variant,
                // not ParamArray - the call site here passes the Array(...) result as a single
                // argument, which a ParamArray would incorrectly nest rather than spread. "values"
                // alternates "name" string literal / bare value pairs, in the same
                // parameter-then-locals order the procedure declares them. "id" is globally
                // unique across an entire instrumentation run via the NextProbeId chaining
                // mechanism. Array() with zero arguments (a procedure with no parameters or
                // locals) is a legitimate, valid emission, not a bug.
                outputLines.Add($"Agent.Probe {currentId}, Array({BuildProbeArgs(inScope)})");
                probeSites.Add(new ProbeSite(currentId, moduleName, originalLine));
                currentId++;
            }

            // Even a line whose probe was suppressed (skip/isBlank) still lexically declares
            // its variables for the rest of the procedure, so this runs unconditionally.
            var declaredHere = VbaParser.ParseDeclarationLine(stripped, VbaParser.ProcedureDeclarationPattern, SymbolKind.Local);
            if (declaredHere.Count > 0)
            {
                inScope.AddRange(declaredHere);
            }

            if (isSelectCaseLine)
            {
                suppressNextProbe = true;
            }

            outputLines.Add(line.Text);
            previousEmittedEndsInContinuationMarker = line.Text.TrimEnd().EndsWith(" _");
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

    private static string BuildProbeArgs(IReadOnlyList<Symbol> symbols)
    {
        var parts = new List<string>();
        foreach (var symbol in symbols)
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
