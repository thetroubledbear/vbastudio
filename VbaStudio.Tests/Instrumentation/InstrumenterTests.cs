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
                        "Agent.Probe 1, Array(\"a\", a)\r\n" +
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

        var lines = result.InstrumentedSource.Split(new[] { "\r\n" }, System.StringSplitOptions.None);
        // Probe before "Dim total As Long" - only the parameters are in scope so far.
        Assert.Equal("Agent.Probe 1, Array(\"x\", x, \"y\", y)", lines[1]);
        // Probe before "Dim label As String" - total's Dim has been passed, label's has not.
        Assert.Equal("Agent.Probe 2, Array(\"x\", x, \"y\", y, \"total\", total)", lines[3]);
        // Probe before "total = x" - both Dim lines have been passed, so all four are in scope,
        // still in the same parameters-then-locals-in-declaration-order that gives this test its name.
        Assert.Equal("Agent.Probe 3, Array(\"x\", x, \"y\", y, \"total\", total, \"label\", label)", lines[5]);
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

        // "y = 2" survives as real, uncommented, separately-executing code - the original
        // continuation-joining bug (which would have merged it into the comment and dropped
        // it from execution) is still fixed. It gets no probe of its own, though: placing one
        // immediately after a line trimming to " _" risks that probe call itself being
        // swallowed if VBA treats the trailing " _" as continuing the comment onto the next
        // physical line - safe under either reading of that open question.
        Assert.Single(result.ProbeSites);
        Assert.Equal(2, result.ProbeSites[0].OriginalLine);
        Assert.Contains("x = 1 ' comment ending in _\r\n    y = 2", result.InstrumentedSource);
    }

    [Fact]
    public void Instrument_LinesInsideWithBlock_GetNoProbe()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    Dim total As Long\r\n" +
                      "    With Sheet1\r\n" +
                      "        total = 1\r\n" +
                      "    End With\r\n" +
                      "    total = 2\r\n" +
                      "End Sub\r\n";
        var module = VbaParser.ParseModule(source, "modWork");
        var procedure = module.Procedures.Single();

        var result = Instrumenter.Instrument(source, procedure, "modWork");

        Assert.Equal(2, result.ProbeSites.Count);
        Assert.Equal(2, result.ProbeSites[0].OriginalLine);
        Assert.Equal(6, result.ProbeSites[1].OriginalLine);
    }

    [Fact]
    public void Instrument_NestedWithBlocks_AllLevelsSkipped()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    With Sheet1\r\n" +
                      "        With Sheet1.Range(\"A1\")\r\n" +
                      "            Value = 1\r\n" +
                      "        End With\r\n" +
                      "        Cells(1, 1).Value = 2\r\n" +
                      "    End With\r\n" +
                      "    Dim done As Boolean\r\n" +
                      "    done = True\r\n" +
                      "End Sub\r\n";
        var module = VbaParser.ParseModule(source, "modWork");
        var procedure = module.Procedures.Single();

        var result = Instrumenter.Instrument(source, procedure, "modWork");

        Assert.Equal(2, result.ProbeSites.Count);
        Assert.Equal(8, result.ProbeSites[0].OriginalLine);
        Assert.Equal(9, result.ProbeSites[1].OriginalLine);
    }

    [Fact]
    public void Instrument_UnterminatedWith_SkipsToEndOfProcedureWithoutThrowing()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    With Sheet1\r\n" +
                      "        Dim x As Long\r\n" +
                      "        x = 1\r\n" +
                      "End Sub\r\n";
        var module = VbaParser.ParseModule(source, "modWork");
        var procedure = module.Procedures.Single();

        var result = Instrumenter.Instrument(source, procedure, "modWork");

        Assert.Empty(result.ProbeSites);
    }

    [Fact]
    public void Instrument_WithAndEndWithLinesThemselves_NeverGetAProbe()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    With Sheet1\r\n" +
                      "        Dim x As Long\r\n" +
                      "    End With\r\n" +
                      "End Sub\r\n";
        var module = VbaParser.ParseModule(source, "modWork");
        var procedure = module.Procedures.Single();

        var result = Instrumenter.Instrument(source, procedure, "modWork");

        Assert.Empty(result.ProbeSites);
        Assert.DoesNotContain("Agent.Probe", result.InstrumentedSource);
    }

    [Fact]
    public void Instrument_SelectCase_NoProbeBetweenSelectCaseAndFirstCase()
    {
        var source = "Public Sub DoWork(x As Long)\r\n" +
                      "    Select Case x\r\n" +
                      "        Case 1\r\n" +
                      "            x = 2\r\n" +
                      "    End Select\r\n" +
                      "End Sub\r\n";
        var module = VbaParser.ParseModule(source, "modWork");
        var procedure = module.Procedures.Single();

        var result = Instrumenter.Instrument(source, procedure, "modWork");

        Assert.Equal(3, result.ProbeSites.Count);
        Assert.Equal(new[] { 2, 4, 5 }, result.ProbeSites.Select(p => p.OriginalLine));
        // No probe line was inserted between "Select Case x" and "Case 1" - they remain adjacent.
        Assert.Contains("    Select Case x\r\n        Case 1\r\n", result.InstrumentedSource);
    }

    [Fact]
    public void Instrument_SelectCaseLineCommentEndingInUnderscore_DoesNotLeakSuppressionToLaterLine()
    {
        var source = "Public Sub DoWork(x As Long)\r\n" +
                      "    Select Case x ' fall-through note _\r\n" +
                      "        Case 1\r\n" +
                      "            x = 2\r\n" +
                      "    End Select\r\n" +
                      "End Sub\r\n";
        var module = VbaParser.ParseModule(source, "modWork");
        var procedure = module.Procedures.Single();

        var result = Instrumenter.Instrument(source, procedure, "modWork");

        // Probed: "Select Case x ..." (line 2), "x = 2" (line 4), "End Select" (line 5).
        // NOT probed: "Case 1" (line 3, suppressed by the Select Case guard).
        // Before this fix, the continuation-marker guard firing on "Case 1" prevented
        // suppressNextProbe from ever resetting, so it leaked forward and incorrectly
        // suppressed "x = 2" too - this pins that it no longer does.
        Assert.Equal(3, result.ProbeSites.Count);
        Assert.Equal(new[] { 2, 4, 5 }, result.ProbeSites.Select(p => p.OriginalLine));
    }

    [Fact]
    public void Instrument_ProbeBeforeDeclaration_DoesNotReferenceNotYetDeclaredLocal()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    Dim x As Long\r\n" +
                      "    Dim y As Long\r\n" +
                      "    x = 1\r\n" +
                      "End Sub\r\n";
        var module = VbaParser.ParseModule(source, "modWork");
        var procedure = module.Procedures.Single();

        var result = Instrumenter.Instrument(source, procedure, "modWork");

        var lines = result.InstrumentedSource.Split(new[] { "\r\n" }, System.StringSplitOptions.None);
        // Probe before "Dim x As Long" - nothing declared yet.
        Assert.Equal("Agent.Probe 1, Array()", lines[1]);
        // Probe before "Dim y As Long" - only x has been declared so far.
        Assert.Equal("Agent.Probe 2, Array(\"x\", x)", lines[3]);
        // Probe before "x = 1" - both x and y are now declared.
        Assert.Equal("Agent.Probe 3, Array(\"x\", x, \"y\", y)", lines[5]);
    }
}
