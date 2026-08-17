using VbaStudio.Core.Parsing;
using Xunit;

namespace VbaStudio.Tests.Parsing;

public class CommentStripperTests
{
    [Fact]
    public void StripComment_NoComment_ReturnsLineUnchanged()
    {
        var result = CommentStripper.StripComment("Dim x As Long");

        Assert.Equal("Dim x As Long", result);
    }

    [Fact]
    public void StripComment_TrailingComment_IsRemoved()
    {
        var result = CommentStripper.StripComment("Dim s As String ' comment after real code");

        Assert.Equal("Dim s As String ", result);
    }

    [Fact]
    public void StripComment_WholeLineComment_ReturnsEmpty()
    {
        var result = CommentStripper.StripComment("' Dim x As Long - this is commented out");

        Assert.Equal("", result);
    }

    [Fact]
    public void StripComment_ApostropheInsideStringLiteral_IsNotTreatedAsCommentStart()
    {
        var result = CommentStripper.StripComment("x = \"a'b\" ' real trailing comment");

        Assert.Equal("x = \"a'b\" ", result);
    }

    [Fact]
    public void StripComment_DoubledQuoteEscapeInsideString_DoesNotConfuseParity()
    {
        // "He said ""hi""" is one string literal containing an escaped quote pair.
        var result = CommentStripper.StripComment("Dim s As String: s = \"He said \"\"hi\"\"\" ' trailing");

        Assert.Equal("Dim s As String: s = \"He said \"\"hi\"\"\" ", result);
    }

    [Fact]
    public void StripComment_EmptyLine_ReturnsEmpty()
    {
        var result = CommentStripper.StripComment("");

        Assert.Equal("", result);
    }
}
