// VbaStudio.Tests/Parsing/VbaParserTests.cs
using System.Linq;
using VbaStudio.Core.Parsing;
using Xunit;

namespace VbaStudio.Tests.Parsing;

public class VbaParserTests
{
    [Fact]
    public void ParseModule_SubWithNoParams_DetectsBoundary()
    {
        var source = "Option Explicit\r\n" +
                      "\r\n" +
                      "Public Sub DoWork()\r\n" +
                      "    Dim x As Long\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        Assert.Equal("modWork", result.ModuleName);
        var proc = Assert.Single(result.Procedures);
        Assert.Equal("DoWork", proc.Name);
        Assert.Equal(ProcedureKind.Sub, proc.Kind);
        Assert.Equal(3, proc.StartLine);
        Assert.Equal(5, proc.EndLine);
    }

    [Fact]
    public void ParseModule_FunctionWithReturnType_DetectsBoundary()
    {
        var source = "Public Function Compute() As Long\r\n" +
                      "    Compute = 42\r\n" +
                      "End Function\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var proc = Assert.Single(result.Procedures);
        Assert.Equal("Compute", proc.Name);
        Assert.Equal(ProcedureKind.Function, proc.Kind);
        Assert.Equal(1, proc.StartLine);
        Assert.Equal(3, proc.EndLine);
    }

    [Theory]
    [InlineData("Property Get Value() As Long", "End Property", ProcedureKind.PropertyGet)]
    [InlineData("Property Let Value(v As Long)", "End Property", ProcedureKind.PropertyLet)]
    [InlineData("Property Set Value(v As Object)", "End Property", ProcedureKind.PropertySet)]
    public void ParseModule_PropertyAccessors_DetectCorrectKind(string header, string footer, ProcedureKind expectedKind)
    {
        var source = $"Public {header}\r\nEnd Property\r\n".Replace("End Property\r\nEnd Property", footer + "\r\n");

        var result = VbaParser.ParseModule(source, "clsWork");

        var proc = Assert.Single(result.Procedures);
        Assert.Equal("Value", proc.Name);
        Assert.Equal(expectedKind, proc.Kind);
    }

    [Fact]
    public void ParseModule_PropertyGetAndLetSharingName_ProduceTwoDistinctEntries()
    {
        var source = "Public Property Get Total() As Long\r\n" +
                      "    Total = 5\r\n" +
                      "End Property\r\n" +
                      "\r\n" +
                      "Public Property Let Total(v As Long)\r\n" +
                      "End Property\r\n";

        var result = VbaParser.ParseModule(source, "clsWork");

        Assert.Equal(2, result.Procedures.Count);
        var get = result.Procedures.Single(p => p.Kind == ProcedureKind.PropertyGet);
        var let = result.Procedures.Single(p => p.Kind == ProcedureKind.PropertyLet);
        Assert.Equal("Total", get.Name);
        Assert.Equal("Total", let.Name);
        Assert.Equal(1, get.StartLine);
        Assert.Equal(3, get.EndLine);
        Assert.Equal(5, let.StartLine);
        Assert.Equal(6, let.EndLine);
    }

    [Fact]
    public void ParseModule_NoExplicitVisibility_DefaultsToDetectedAsPublicKindStill()
    {
        // Missing Public/Private defaults to Public per VBA's own rule; the parser does not
        // surface a separate "implicit vs explicit" flag - this just confirms the boundary is
        // still detected without an explicit visibility keyword.
        var source = "Sub QuickHelper()\r\nEnd Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var proc = Assert.Single(result.Procedures);
        Assert.Equal("QuickHelper", proc.Name);
    }

    [Fact]
    public void ParseModule_MultipleProcedures_AllDetectedWithCorrectRanges()
    {
        var source = "Public Sub First()\r\n" +
                      "End Sub\r\n" +
                      "\r\n" +
                      "Public Sub Second()\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        Assert.Equal(2, result.Procedures.Count);
        Assert.Equal("First", result.Procedures[0].Name);
        Assert.Equal(1, result.Procedures[0].StartLine);
        Assert.Equal(2, result.Procedures[0].EndLine);
        Assert.Equal("Second", result.Procedures[1].Name);
        Assert.Equal(4, result.Procedures[1].StartLine);
        Assert.Equal(5, result.Procedures[1].EndLine);
    }

    [Fact]
    public void ParseModule_ContinuedSignature_ReportsPhysicalNotJoinedLineNumbers()
    {
        var source = "Public Sub DoWork(a As Long, _\r\n" +
                      "    b As String)\r\n" +
                      "    Dim total As Long\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var proc = Assert.Single(result.Procedures);
        Assert.Equal(1, proc.StartLine);
        Assert.Equal(4, proc.EndLine);
    }

    [Fact]
    public void ParseModule_UnterminatedSub_ClosesAtEndOfFileWithoutThrowing()
    {
        var source = "Public Sub Broken()\r\n" +
                      "    Dim x As Long\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var proc = Assert.Single(result.Procedures);
        Assert.Equal("Broken", proc.Name);
        Assert.Equal(1, proc.StartLine);
        Assert.Equal(2, proc.EndLine);
    }

    [Fact]
    public void ParseModule_CommentedOutProcedureSignature_IsNotDetectedAsABoundary()
    {
        var source = "' Public Sub Disabled()\r\n" +
                      "' End Sub\r\n" +
                      "Public Sub Real()\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var proc = Assert.Single(result.Procedures);
        Assert.Equal("Real", proc.Name);
    }

    [Fact]
    public void ParseModule_NoProcedures_ReturnsEmptyNotNull()
    {
        var source = "Option Explicit\r\n";

        var result = VbaParser.ParseModule(source, "modEmpty");

        Assert.Equal("modEmpty", result.ModuleName);
        Assert.Empty(result.Procedures);
        Assert.Empty(result.ModuleVariables);
    }

    [Fact]
    public void ParseModule_EveryProcedure_HasEmptyParametersAndLocalsForNow()
    {
        var source = "Public Sub DoWork(a As Long)\r\n" +
                      "    Dim x As Long\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var proc = Assert.Single(result.Procedures);
        Assert.Empty(proc.Parameters);
        Assert.Empty(proc.Locals);
    }
}
