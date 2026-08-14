// VbaStudio.Core/Parsing/VbaParser.cs
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Text.RegularExpressions;
using VbaStudio.Core.Model;

namespace VbaStudio.Core.Parsing;

public static class VbaParser
{
    private static readonly Regex ProcedureHeaderPattern = new(
        @"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?<kind>Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+(?<name>\w+)\s*\((?<params>(?:[^()]|\(\))*)\)(?:\s+As\s+[\w.]+(?:\(\))?)?\s*$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ProcedureEndPattern = new(
        @"^\s*End\s+(?:Sub|Function|Property)\s*$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static ModuleSymbols ParseModule(string sourceText, string moduleName)
    {
        // Comments must be stripped from each physical line BEFORE continuation-joining, not after:
        // VBA disallows continuing a comment across lines, so a comment whose text happens to end in
        // " _" must not be allowed to make LineJoiner swallow the next physical line into it.
        var physicalLines = SplitPhysicalLines(sourceText);
        var strippedLines = physicalLines.Select(CommentStripper.StripComment).ToList();
        var joined = LineJoiner.Join(strippedLines);

        var procedures = new List<ProcedureSymbols>();

        // Tracks, per joined line, whether it falls inside a detected procedure's range. Module-level
        // declaration scanning below is defined as the set-complement of this array ("every line NOT
        // inside a procedure"), not "every line before the first procedure" — so declarations that
        // appear between or after procedures are still picked up correctly.
        var inProcedureLine = new bool[joined.Count];
        var i = 0;

        while (i < joined.Count)
        {
            var headerMatch = ProcedureHeaderPattern.Match(joined[i].Text);
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
                if (ProcedureEndPattern.IsMatch(joined[j].Text))
                {
                    break;
                }

                j++;
            }

            var locals = new List<Symbol>();
            for (var bodyIndex = bodyStart; bodyIndex < j && bodyIndex < joined.Count; bodyIndex++)
            {
                locals.AddRange(ParseDeclarationLine(joined[bodyIndex].Text, ProcedureDeclarationPattern, SymbolKind.Local));
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

            moduleVariables.AddRange(ParseModuleDeclarationLine(joined[lineIndex].Text));
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

            // VBA disallows ByRef/ByVal on ParamArray, so PassingMode is forced to null rather than
            // defaulting to "ByRef" for it.
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

    internal static readonly Regex ProcedureDeclarationPattern = new(
        @"^\s*(?<kw>Dim|Static|Const)\s+(?<rest>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // These four module-level patterns are tried in this exact order in ParseModuleDeclarationLine:
    // Const-with-visibility and bare Const must be checked before the bare-visibility fallback, or a
    // line like "Private Const Max = 100" would be misclassified as a plain ModuleVariable with a
    // garbled "rest" (the word "Const" itself would end up captured as part of the declaration text).
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

    // Optional "WithEvents" precedes the name (module-level only, but harmless for locals);
    // optional "New" precedes the type name for "As New Type" declarations — neither is captured.
    private static readonly Regex DeclaredVariablePattern = new(
        @"^\s*(?:WithEvents\s+)?(?<name>\w+)\s*(?<array>\(\s*[\w,\s]*\s*\))?\s*(?:As\s+(?:New\s+)?(?<type>[\w.]+))?\s*(?:=.*)?$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    internal static IReadOnlyList<Symbol> ParseDeclarationLine(string line, Regex declarationPattern, SymbolKind nonConstKind)
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

    // Paren-depth-aware because a declaration or parameter list can contain array-size parens
    // (e.g. "arr(10, 20)") whose internal commas must not be treated as top-level separators.
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

    // Same four kinds SyncEngine's own AllKinds writes back to disk — keeps folder/extension/encoding
    // handling for reading a project in lockstep with how SyncEngine writes one.
    private static readonly ModuleKind[] AllKinds =
    {
        ModuleKind.Standard, ModuleKind.Class, ModuleKind.UserForm, ModuleKind.Document
    };

    public static IReadOnlyList<ModuleSymbols> ParseProject(IFileSystem fileSystem, string srcDir)
    {
        var results = new List<ModuleSymbols>();

        foreach (var kind in AllKinds)
        {
            var folder = fileSystem.Path.Combine(srcDir, kind.SourceFolder());
            if (!fileSystem.Directory.Exists(folder))
            {
                continue;
            }

            var extension = kind.FileExtension();
            foreach (var path in fileSystem.Directory.EnumerateFiles(folder))
            {
                if (!fileSystem.Path.GetExtension(path).Equals(extension, System.StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var moduleName = fileSystem.Path.GetFileNameWithoutExtension(path);
                var sourceText = fileSystem.File.ReadAllText(path, kind.SourceEncoding());
                results.Add(ParseModule(sourceText, moduleName));
            }
        }

        return results;
    }
}
