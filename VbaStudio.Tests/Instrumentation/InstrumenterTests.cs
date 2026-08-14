// VbaStudio.Tests/Instrumentation/InstrumenterTests.cs
using System.Linq;
using System.Text.RegularExpressions;
using VbaStudio.Core.Instrumentation;
using VbaStudio.Core.Parsing;
using Xunit;

namespace VbaStudio.Tests.Instrumentation;

public class InstrumenterTests
{
    [Fact]
    public void Instrument_SimpleProcedure_ProducesExactExpectedShape()
    {
        var source = "Public Sub DoWork(a As Long)\r\n" +
                      "    Dim total As Long\r\n" +
                      "    total = a + 1\r\n" +
                      "End Sub\r\n";
        var module = VbaParser.ParseModule(source, "modWork");
        var procedure = module.Procedures.Single();

        var result = Instrumenter.Instrument(source, procedure, "modWork");

        var expected = "Public Sub DoWork(a As Long)\r\n" +
                        "Agent.Probe 1, Array(\"a\", a, \"total\", total)\r\n" +
                        "    Dim total As Long\r\n" +
                        "Agent.Probe 2, Array(\"a\", a, \"total\", total)\r\n" +
                        "    total = a + 1\r\n" +
                        "End Sub";
        Assert.Equal(expected, result.InstrumentedSource);
    }

    [Fact]
    public void Instrument_ProbeSites_MapEachProbeIdToOriginalPhysicalLine()
    {
        var source = "Public Sub DoWork(a As Long)\r\n" +
                      "    Dim total As Long\r\n" +
                      "    total = a + 1\r\n" +
                      "End Sub\r\n";
        var module = VbaParser.ParseModule(source, "modWork");
        var procedure = module.Procedures.Single();

        var result = Instrumenter.Instrument(source, procedure, "modWork");

        Assert.Equal(2, result.ProbeSites.Count);
        Assert.Equal(new ProbeSite(1, "modWork", 2), result.ProbeSites[0]);
        Assert.Equal(new ProbeSite(2, "modWork", 3), result.ProbeSites[1]);
    }

    [Fact]
    public void Instrument_ArrayArgs_ListParametersThenLocalsInDeclarationOrder()
    {
        var source = "Public Sub DoWork(x As Long, y As String)\r\n" +
                      "    Dim total As Long\r\n" +
                      "    Dim label As String\r\n" +
                      "    total = x\r\n" +
                      "End Sub\r\n";
        var module = VbaParser.ParseModule(source, "modWork");
        var procedure = module.Procedures.Single();

        var result = Instrumenter.Instrument(source, procedure, "modWork");

        Assert.Contains("Array(\"x\", x, \"y\", y, \"total\", total, \"label\", label)", result.InstrumentedSource);
    }

    [Fact]
    public void Instrument_BlankAndCommentOnlyLines_GetNoProbe()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    Dim x As Long\r\n" +
                      "\r\n" +
                      "    ' just a comment\r\n" +
                      "    x = 1\r\n" +
                      "End Sub\r\n";
        var module = VbaParser.ParseModule(source, "modWork");
        var procedure = module.Procedures.Single();

        var result = Instrumenter.Instrument(source, procedure, "modWork");

        Assert.Equal(2, result.ProbeSites.Count);
        Assert.Equal(2, result.ProbeSites[0].OriginalLine);
        Assert.Equal(5, result.ProbeSites[1].OriginalLine);
    }

    [Fact]
    public void Instrument_ContinuedStatement_ProbedOnceNotSplit()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    Dim total As Long\r\n" +
                      "    total = 1 + _\r\n" +
                      "        2\r\n" +
                      "End Sub\r\n";
        var module = VbaParser.ParseModule(source, "modWork");
        var procedure = module.Procedures.Single();

        var result = Instrumenter.Instrument(source, procedure, "modWork");

        Assert.Equal(2, result.ProbeSites.Count);
        Assert.Equal(2, result.ProbeSites[0].OriginalLine);
        Assert.Equal(3, result.ProbeSites[1].OriginalLine);
        var probeCount = Regex.Matches(result.InstrumentedSource, "Agent.Probe").Count;
        Assert.Equal(2, probeCount);
    }

    [Fact]
    public void Instrument_SingleLineIf_ProbedAsOneLineNotSplit()
    {
        var source = "Public Sub DoWork(x As Long)\r\n" +
                      "    If x > 0 Then x = x - 1\r\n" +
                      "End Sub\r\n";
        var module = VbaParser.ParseModule(source, "modWork");
        var procedure = module.Procedures.Single();

        var result = Instrumenter.Instrument(source, procedure, "modWork");

        Assert.Single(result.ProbeSites);
        Assert.Equal(2, result.ProbeSites[0].OriginalLine);
        var probeCount = Regex.Matches(result.InstrumentedSource, "Agent.Probe").Count;
        Assert.Equal(1, probeCount);
    }

    [Fact]
    public void Instrument_StartingProbeId_SequencesFromCallerSuppliedValue()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    Dim x As Long\r\n" +
                      "    Dim y As Long\r\n" +
                      "End Sub\r\n";
        var module = VbaParser.ParseModule(source, "modWork");
        var procedure = module.Procedures.Single();

        var result = Instrumenter.Instrument(source, procedure, "modWork", startingProbeId: 50);

        Assert.Equal(50, result.ProbeSites[0].ProbeId);
        Assert.Equal(51, result.ProbeSites[1].ProbeId);
        Assert.Equal(52, result.NextProbeId);
    }

    [Fact]
    public void Instrument_EmptyBodyProcedure_YieldsZeroProbes()
    {
        var source = "Public Sub DoNothing()\r\nEnd Sub\r\n";
        var module = VbaParser.ParseModule(source, "modWork");
        var procedure = module.Procedures.Single();

        var result = Instrumenter.Instrument(source, procedure, "modWork");

        Assert.Empty(result.ProbeSites);
        Assert.Equal(1, result.NextProbeId);
        Assert.Equal("Public Sub DoNothing()\r\nEnd Sub", result.InstrumentedSource);
    }

    [Fact]
    public void Instrument_LinesOutsideTargetProcedure_PassThroughUnchanged()
    {
        var source = "Public Sub Before()\r\n" +
                      "    Dim ignored As Long\r\n" +
                      "End Sub\r\n" +
                      "\r\n" +
                      "Public Sub DoWork()\r\n" +
                      "    Dim x As Long\r\n" +
                      "End Sub\r\n" +
                      "\r\n" +
                      "Public Sub After()\r\n" +
                      "    Dim alsoIgnored As Long\r\n" +
                      "End Sub\r\n";
        var module = VbaParser.ParseModule(source, "modWork");
        var procedure = module.Procedures.Single(p => p.Name == "DoWork");

        var result = Instrumenter.Instrument(source, procedure, "modWork");

        Assert.Contains("Public Sub Before()\r\n    Dim ignored As Long\r\nEnd Sub", result.InstrumentedSource);
        Assert.Contains("Public Sub After()\r\n    Dim alsoIgnored As Long\r\nEnd Sub", result.InstrumentedSource);
        Assert.DoesNotContain("ignored As Long\r\nAgent.Probe", result.InstrumentedSource);
        Assert.DoesNotContain("alsoIgnored As Long\r\nAgent.Probe", result.InstrumentedSource);
    }

    [Fact]
    public void Instrument_CommentEndingInUnderscore_DoesNotFalselyTriggerContinuation()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    x = 1 ' comment ending in _\r\n" +
                      "    y = 2\r\n" +
                      "End Sub\r\n";
        var module = VbaParser.ParseModule(source, "modWork");
        var procedure = module.Procedures.Single();

        var result = Instrumenter.Instrument(source, procedure, "modWork");

        // Two independent statements, two probes - not merged into one.
        Assert.Equal(2, result.ProbeSites.Count);
        Assert.Equal(2, result.ProbeSites[0].OriginalLine);
        Assert.Equal(3, result.ProbeSites[1].OriginalLine);
        // The real comment survives in the output, unmodified.
        Assert.Contains("x = 1 ' comment ending in _", result.InstrumentedSource);
        // "y = 2" must NOT have been swallowed into the comment - it must appear as its own
        // real (uncommented) statement, immediately preceded by its own probe line.
        Assert.Contains("Agent.Probe 2, Array()\r\n    y = 2", result.InstrumentedSource);
    }
}
