namespace VbaStudio.Core.Parsing;

internal static class CommentStripper
{
    public static string StripComment(string line)
    {
        var inString = false;

        for (var i = 0; i < line.Length; i++)
        {
            var c = line[i];
            if (c == '"')
            {
                inString = !inString;
            }
            else if (c == '\'' && !inString)
            {
                return line.Substring(0, i);
            }
        }

        return line;
    }
}
