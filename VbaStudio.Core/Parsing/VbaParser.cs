// VbaStudio.Core/Parsing/VbaParser.cs
using System.Collections.Generic;
using System.IO.Abstractions;
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
        var inProcedureLine = new bool[joined.Count];
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
            var bodyStart = i + 1;
            var j = bodyStart;
            while (j < joined.Count)
            {
                endLine = joined[j].EndPhysicalLine;
                if (ProcedureEndPattern.IsMatch(clean[j]))
                {
                    break;
                }

                j++;
            }

            var locals = new List<Symbol>();
            for (var bodyIndex = bodyStart; bodyIndex < j && bodyIndex < joined.Count; bodyIndex++)
            {
                locals.AddRange(ParseDeclarationLine(clean[bodyIndex], ProcedureDeclarationPattern, SymbolKind.Local));
            }

            for (var markIndex = i; markIndex <= j && markIndex < joined.Count; markIndex++)
            {
                inProcedureLine[markIndex] = true;
            }

            procedures.Add(new ProcedureSymbols(
                name,
                kind,
                startLine,
                endLine,
                Parameters: ParseParameters(headerMatch.Groups["params"].Value),
                Locals: locals));

            i = j + 1;
        }

        var moduleVariables = new List<Symbol>();
        for (var lineIndex = 0; lineIndex < joined.Count; lineIndex++)
        {
            if (inProcedureLine[lineIndex])
            {
                continue;
            }

            moduleVariables.AddRange(ParseModuleDeclarationLine(clean[lineIndex]));
        }

        return new ModuleSymbols(moduleName, moduleVariables, procedures);
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

    private static readonly Regex ParameterPattern = new(
        @"^\s*(?:(?<optional>Optional)\s+)?(?:(?<passing>ByRef|ByVal)\s+)?(?:(?<paramarray>ParamArray)\s+)?(?<name>\w+)\s*(?<array>\(\s*\))?\s*(?:As\s+(?<type>[\w.]+))?\s*(?:=.*)?$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static IReadOnlyList<Symbol> ParseParameters(string rawParams)
    {
        if (string.IsNullOrWhiteSpace(rawParams))
        {
            return System.Array.Empty<Symbol>();
        }

        var results = new List<Symbol>();
        foreach (var segment in SplitTopLevel(rawParams, ','))
        {
            var trimmed = segment.Trim();
            if (trimmed.Length == 0)
            {
                continue;
            }

            var match = ParameterPattern.Match(trimmed);
            if (!match.Success)
            {
                continue;
            }

            var isParamArray = match.Groups["paramarray"].Success;
            var isArray = match.Groups["array"].Success || isParamArray;
            var declaredType = match.Groups["type"].Success ? match.Groups["type"].Value : "Variant";
            string? passingMode = isParamArray
                ? null
                : match.Groups["passing"].Success
                    ? NormalizeKeyword(match.Groups["passing"].Value)
                    : "ByRef";

            results.Add(new Symbol(
                match.Groups["name"].Value,
                declaredType,
                SymbolKind.Parameter,
                isArray,
                IsOptional: match.Groups["optional"].Success,
                passingMode));
        }

        return results;
    }

    private static string NormalizeKeyword(string raw) =>
        raw.Equals("ByVal", System.StringComparison.OrdinalIgnoreCase) ? "ByVal" : "ByRef";

    private static readonly Regex ProcedureDeclarationPattern = new(
        @"^\s*(?<kw>Dim|Static|Const)\s+(?<rest>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ModuleConstWithVisibilityPattern = new(
        @"^\s*(?:Public|Private|Global)\s+Const\s+(?<rest>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ModuleConstBarePattern = new(
        @"^\s*Const\s+(?<rest>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ModuleDimPattern = new(
        @"^\s*Dim\s+(?<rest>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ModuleVisibilityOnlyPattern = new(
        @"^\s*(?:Public|Private|Global)\s+(?<rest>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static IReadOnlyList<Symbol> ParseModuleDeclarationLine(string line)
    {
        var constWithVisibility = ModuleConstWithVisibilityPattern.Match(line);
        if (constWithVisibility.Success)
        {
            return ParseDeclaredVariables(constWithVisibility.Groups["rest"].Value, SymbolKind.Const);
        }

        var constBare = ModuleConstBarePattern.Match(line);
        if (constBare.Success)
        {
            return ParseDeclaredVariables(constBare.Groups["rest"].Value, SymbolKind.Const);
        }

        var dim = ModuleDimPattern.Match(line);
        if (dim.Success)
        {
            return ParseDeclaredVariables(dim.Groups["rest"].Value, SymbolKind.ModuleVariable);
        }

        var visibilityOnly = ModuleVisibilityOnlyPattern.Match(line);
        if (visibilityOnly.Success)
        {
            return ParseDeclaredVariables(visibilityOnly.Groups["rest"].Value, SymbolKind.ModuleVariable);
        }

        return System.Array.Empty<Symbol>();
    }

    private static readonly Regex DeclaredVariablePattern = new(
        @"^\s*(?<name>\w+)\s*(?<array>\(\s*[\w,\s]*\s*\))?\s*(?:As\s+(?<type>[\w.]+))?\s*(?:=.*)?$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static IReadOnlyList<Symbol> ParseDeclarationLine(string line, Regex declarationPattern, SymbolKind nonConstKind)
    {
        var match = declarationPattern.Match(line);
        if (!match.Success)
        {
            return System.Array.Empty<Symbol>();
        }

        var kw = match.Groups["kw"].Value;
        var kind = kw.Equals("Const", System.StringComparison.OrdinalIgnoreCase) ? SymbolKind.Const : nonConstKind;

        return ParseDeclaredVariables(match.Groups["rest"].Value, kind);
    }

    private static IReadOnlyList<Symbol> ParseDeclaredVariables(string rawDeclarations, SymbolKind kind)
    {
        var results = new List<Symbol>();
        foreach (var segment in SplitTopLevel(rawDeclarations, ','))
        {
            var trimmed = segment.Trim();
            if (trimmed.Length == 0)
            {
                continue;
            }

            var match = DeclaredVariablePattern.Match(trimmed);
            if (!match.Success)
            {
                continue;
            }

            var declaredType = match.Groups["type"].Success ? match.Groups["type"].Value : "Variant";

            results.Add(new Symbol(
                match.Groups["name"].Value,
                declaredType,
                kind,
                IsArray: match.Groups["array"].Success,
                IsOptional: false,
                PassingMode: null));
        }

        return results;
    }

    private static IReadOnlyList<string> SplitTopLevel(string text, char separator)
    {
        var parts = new List<string>();
        var depth = 0;
        var start = 0;

        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];
            if (c == '(')
            {
                depth++;
            }
            else if (c == ')')
            {
                depth--;
            }
            else if (c == separator && depth == 0)
            {
                parts.Add(text.Substring(start, i - start));
                start = i + 1;
            }
        }

        parts.Add(text.Substring(start));
        return parts;
    }

    private static readonly string[] SourceSubfolders = { "Modules", "Classes", "Forms", "Sheets" };

    public static IReadOnlyList<ModuleSymbols> ParseProject(IFileSystem fileSystem, string srcDir)
    {
        var results = new List<ModuleSymbols>();

        foreach (var subfolder in SourceSubfolders)
        {
            var folder = fileSystem.Path.Combine(srcDir, subfolder);
            if (!fileSystem.Directory.Exists(folder))
            {
                continue;
            }

            foreach (var path in fileSystem.Directory.EnumerateFiles(folder))
            {
                var moduleName = fileSystem.Path.GetFileNameWithoutExtension(path);
                var sourceText = fileSystem.File.ReadAllText(path);
                results.Add(ParseModule(sourceText, moduleName));
            }
        }

        return results;
    }
}
