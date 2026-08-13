using System;
using System.Linq;

namespace VbaStudio.Core.Excel;

public static class VbaSourceText
{
    public static string[] RemoveHeaderAttributeLines(string[] lines)
    {
        int nonAttributeLines = lines
            .TakeWhile(line => !line.StartsWith("Attribute", StringComparison.Ordinal))
            .Count();
        int attributeLines = lines.Skip(nonAttributeLines)
            .TakeWhile(line => line.StartsWith("Attribute", StringComparison.Ordinal))
            .Count();
        int declarationsStartLine = nonAttributeLines + attributeLines + 1;

        return lines.Skip(declarationsStartLine - 1).ToArray();
    }

    public static string[] TrimExtraLeadingBlankLine(string[] originalContentLines, string[] exportedLines)
    {
        int legitEmptyLineCount = originalContentLines.TakeWhile(string.IsNullOrWhiteSpace).Count();

        int nonAttributeLines = exportedLines
            .TakeWhile(line => !line.StartsWith("Attribute", StringComparison.Ordinal))
            .Count();
        int attributeLines = exportedLines.Skip(nonAttributeLines)
            .TakeWhile(line => line.StartsWith("Attribute", StringComparison.Ordinal))
            .Count();
        int declarationsStartLine = nonAttributeLines + attributeLines + 1;

        int emptyLineCount = exportedLines.Skip(declarationsStartLine - 1)
            .TakeWhile(string.IsNullOrWhiteSpace)
            .Count();

        if (emptyLineCount <= legitEmptyLineCount)
        {
            return exportedLines;
        }

        int extra = emptyLineCount - legitEmptyLineCount;
        return exportedLines.Take(declarationsStartLine - 1)
            .Concat(exportedLines.Skip(declarationsStartLine - 1 + extra))
            .ToArray();
    }
}
