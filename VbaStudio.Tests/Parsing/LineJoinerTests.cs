using System.Collections.Generic;
using VbaStudio.Core.Parsing;
using Xunit;

namespace VbaStudio.Tests.Parsing;

public class LineJoinerTests
{
    [Fact]
    public void Join_NoContinuations_EachLineStandsAlone()
    {
        var lines = new[] { "Dim x As Long", "Dim y As String" };

        var result = LineJoiner.Join(lines);

        Assert.Equal(2, result.Count);
        Assert.Equal("Dim x As Long", result[0].Text);
        Assert.Equal(1, result[0].StartPhysicalLine);
        Assert.Equal(1, result[0].EndPhysicalLine);
        Assert.Equal("Dim y As String", result[1].Text);
        Assert.Equal(2, result[1].StartPhysicalLine);
        Assert.Equal(2, result[1].EndPhysicalLine);
    }

    [Fact]
    public void Join_SingleContinuation_MergesIntoOneLogicalLine()
    {
        var lines = new[]
        {
            "Public Sub DoWork(a As Long, _",
            "    b As String)",
        };

        var result = LineJoiner.Join(lines);

        var joined = Assert.Single(result);
        Assert.Equal("Public Sub DoWork(a As Long,     b As String)", joined.Text);
        Assert.Equal(1, joined.StartPhysicalLine);
        Assert.Equal(2, joined.EndPhysicalLine);
    }

    [Fact]
    public void Join_ChainOfContinuations_MergesAllIntoOneLogicalLine()
    {
        var lines = new[]
        {
            "Public Sub DoWork(a As Long, _",
            "    b As String, _",
            "    c As Boolean)",
        };

        var result = LineJoiner.Join(lines);

        var joined = Assert.Single(result);
        Assert.Contains("a As Long", joined.Text);
        Assert.Contains("b As String", joined.Text);
        Assert.Contains("c As Boolean", joined.Text);
        Assert.Equal(1, joined.StartPhysicalLine);
        Assert.Equal(3, joined.EndPhysicalLine);
    }

    [Fact]
    public void Join_ContinuationFollowedByNormalLine_ResumesLineByLineAfter()
    {
        var lines = new[]
        {
            "Public Sub DoWork(a As Long, _",
            "    b As String)",
            "    Dim total As Long",
        };

        var result = LineJoiner.Join(lines);

        Assert.Equal(2, result.Count);
        Assert.Equal(1, result[0].StartPhysicalLine);
        Assert.Equal(2, result[0].EndPhysicalLine);
        Assert.Equal("    Dim total As Long", result[1].Text);
        Assert.Equal(3, result[1].StartPhysicalLine);
        Assert.Equal(3, result[1].EndPhysicalLine);
    }

    [Fact]
    public void Join_TrailingUnderscoreWithNoTextAfter_DoesNotRequireASpaceOnlyLine()
    {
        // A line ending " _" with nothing following it in the array (EOF mid-continuation) is
        // malformed/dangling VBA - best-effort: left as-is, not joined with anything, no throw.
        var lines = new[] { "Public Sub DoWork(a As Long, _" };

        var result = LineJoiner.Join(lines);

        var joined = Assert.Single(result);
        Assert.Equal("Public Sub DoWork(a As Long, _", joined.Text);
        Assert.Equal(1, joined.StartPhysicalLine);
        Assert.Equal(1, joined.EndPhysicalLine);
    }

    [Fact]
    public void Join_UnderscoreNotPrecededBySpace_IsNotTreatedAsContinuation()
    {
        // VBA requires the space before the underscore. "MyVar_" is a legal identifier suffix,
        // not a continuation marker.
        var lines = new[] { "Dim MyVar_ As Long", "Dim other As String" };

        var result = LineJoiner.Join(lines);

        Assert.Equal(2, result.Count);
        Assert.Equal("Dim MyVar_ As Long", result[0].Text);
    }

    [Fact]
    public void Join_EmptyInput_ReturnsEmpty()
    {
        var result = LineJoiner.Join(new List<string>());

        Assert.Empty(result);
    }
}
