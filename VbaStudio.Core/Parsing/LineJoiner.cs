using System.Collections.Generic;
using System.Text;

namespace VbaStudio.Core.Parsing;

internal sealed record JoinedLine(string Text, int StartPhysicalLine, int EndPhysicalLine);

internal static class LineJoiner
{
    public static IReadOnlyList<JoinedLine> Join(IReadOnlyList<string> physicalLines)
    {
        var result = new List<JoinedLine>();
        var i = 0;

        while (i < physicalLines.Count)
        {
            var startLine = i + 1;
            var sb = new StringBuilder();
            var current = i;

            while (true)
            {
                var line = physicalLines[current];
                var trimmedEnd = line.TrimEnd();
                var hasMoreLines = current + 1 < physicalLines.Count;

                if (hasMoreLines && trimmedEnd.EndsWith(" _"))
                {
                    sb.Append(trimmedEnd, 0, trimmedEnd.Length - 2);
                    sb.Append(' ');
                    current++;
                    continue;
                }

                sb.Append(line);
                break;
            }

            result.Add(new JoinedLine(sb.ToString(), startLine, current + 1));
            i = current + 1;
        }

        return result;
    }
}
