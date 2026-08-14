// VbaStudio.Core/Parsing/VbaParser.cs
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace VbaStudio.Core.Parsing;

public static class VbaParser
{
    private static readonly Regex ProcedureHeaderPattern = new(
        @"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?<kind>Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+(?<name>\w+)\s*\((?<params>(?:[^()]|\(\))*)\)(?:\s+As\s+\w+(?:\(\))?)?\s*$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ProcedureEndPattern = new(
        @"^\s*End\s+(?:Sub|Function|Property)\s*$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static ModuleSymbols ParseModule(string sourceText, string moduleName)
    {
        var physicalLines = SplitPhysicalLines(sourceText);
        var joined = LineJoiner.Join(physicalLines);
        var clean = joined.Select(j => CommentStripper.StripComment(j.Text)).ToList();

        var procedures = new List<ProcedureSymbols>();
        var i = 0;

        while (i < joined.Count)
        {
            var headerMatch = ProcedureHeaderPattern.Match(clean[i]);
            if (!headerMatch.Success)
            {
                i++;
                continue;
            }

            var kind = ParseProcedureKind(headerMatch.Groups["kind"].Value);
            var name = headerMatch.Groups["name"].Value;
            var startLine = joined[i].StartPhysicalLine;

            var endLine = joined[i].EndPhysicalLine;
            var j = i + 1;
            while (j < joined.Count)
            {
                endLine = joined[j].EndPhysicalLine;
                if (ProcedureEndPattern.IsMatch(clean[j]))
                {
                    break;
                }

                j++;
            }

            procedures.Add(new ProcedureSymbols(
                name,
                kind,
                startLine,
                endLine,
                Parameters: System.Array.Empty<Symbol>(),
                Locals: System.Array.Empty<Symbol>()));

            i = j + 1;
        }

        return new ModuleSymbols(moduleName, System.Array.Empty<Symbol>(), procedures);
    }

    private static ProcedureKind ParseProcedureKind(string rawKind)
    {
        var normalized = Regex.Replace(rawKind, @"\s+", " ").Trim().ToLowerInvariant();
        return normalized switch
        {
            "sub" => ProcedureKind.Sub,
            "function" => ProcedureKind.Function,
            "property get" => ProcedureKind.PropertyGet,
            "property let" => ProcedureKind.PropertyLet,
            "property set" => ProcedureKind.PropertySet,
            _ => ProcedureKind.Sub,
        };
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
