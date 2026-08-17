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
    public void ParseModule_PublicSub_VisibilityIsPublic()
    {
        var source = "Public Sub DoWork()\r\nEnd Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var proc = Assert.Single(result.Procedures);
        Assert.Equal(ProcedureVisibility.Public, proc.Visibility);
    }

    [Fact]
    public void ParseModule_PrivateSub_VisibilityIsPrivate()
    {
        var source = "Private Sub Helper()\r\nEnd Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var proc = Assert.Single(result.Procedures);
        Assert.Equal(ProcedureVisibility.Private, proc.Visibility);
    }

    [Fact]
    public void ParseModule_SubWithNoVisibilityKeyword_DefaultsToPublic()
    {
        var source = "Sub DoWork()\r\nEnd Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var proc = Assert.Single(result.Procedures);
        Assert.Equal(ProcedureVisibility.Public, proc.Visibility);
    }

    [Fact]
    public void ParseModule_FriendFunction_VisibilityIsFriend()
    {
        var source = "Friend Function Calc() As Long\r\n    Calc = 1\r\nEnd Function\r\n";

        var result = VbaParser.ParseModule(source, "clsWork");

        var proc = Assert.Single(result.Procedures);
        Assert.Equal(ProcedureVisibility.Friend, proc.Visibility);
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
    public void ParseModule_ProcedureWithParametersAndLocals_PopulatesBoth()
    {
        var source = "Public Sub DoWork(a As Long)\r\n" +
                      "    Dim x As Long\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var proc = Assert.Single(result.Procedures);
        Assert.NotEmpty(proc.Parameters);
        var local = Assert.Single(proc.Locals);
        Assert.Equal("x", local.Name);
        Assert.Equal("Long", local.DeclaredType);
        Assert.Equal(SymbolKind.Local, local.Kind);
    }

    [Fact]
    public void ParseModule_SimpleParameter_DefaultsToByRefAndVariantWhenUntyped()
    {
        var source = "Public Sub DoWork(x)\r\nEnd Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var param = Assert.Single(result.Procedures.Single().Parameters);
        Assert.Equal("x", param.Name);
        Assert.Equal("Variant", param.DeclaredType);
        Assert.Equal(SymbolKind.Parameter, param.Kind);
        Assert.Equal("ByRef", param.PassingMode);
        Assert.False(param.IsOptional);
        Assert.False(param.IsArray);
    }

    [Fact]
    public void ParseModule_TypedByValParameter_ParsesCorrectly()
    {
        var source = "Public Sub DoWork(ByVal count As Long)\r\nEnd Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var param = Assert.Single(result.Procedures.Single().Parameters);
        Assert.Equal("count", param.Name);
        Assert.Equal("Long", param.DeclaredType);
        Assert.Equal("ByVal", param.PassingMode);
    }

    [Fact]
    public void ParseModule_OptionalParameterWithDefaultValue_SetsIsOptionalTrue()
    {
        var source = "Public Sub DoWork(Optional x As Long = 5)\r\nEnd Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var param = Assert.Single(result.Procedures.Single().Parameters);
        Assert.Equal("x", param.Name);
        Assert.True(param.IsOptional);
        Assert.Equal("Long", param.DeclaredType);
    }

    [Fact]
    public void ParseModule_OptionalParameterWithoutDefaultValue_SetsIsOptionalTrue()
    {
        var source = "Public Sub DoWork(Optional label As String)\r\nEnd Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var param = Assert.Single(result.Procedures.Single().Parameters);
        Assert.True(param.IsOptional);
    }

    [Fact]
    public void ParseModule_ParamArrayParameter_SetsArrayTrueAndPassingModeNull()
    {
        var source = "Public Sub DoWork(ParamArray items() As Variant)\r\nEnd Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var param = Assert.Single(result.Procedures.Single().Parameters);
        Assert.Equal("items", param.Name);
        Assert.True(param.IsArray);
        Assert.Null(param.PassingMode);
        Assert.False(param.IsOptional);
    }

    [Fact]
    public void ParseModule_ArrayParameter_SetsIsArrayTrue()
    {
        var source = "Public Sub DoWork(values() As Long)\r\nEnd Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var param = Assert.Single(result.Procedures.Single().Parameters);
        Assert.Equal("values", param.Name);
        Assert.True(param.IsArray);
        Assert.Equal("Long", param.DeclaredType);
    }

    [Fact]
    public void ParseModule_MultipleParameters_AllParsedInOrder()
    {
        var source = "Public Sub DoWork(a As Long, ByVal b As String, Optional c As Boolean = False)\r\nEnd Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var parameters = result.Procedures.Single().Parameters;
        Assert.Equal(3, parameters.Count);
        Assert.Equal("a", parameters[0].Name);
        Assert.Equal("b", parameters[1].Name);
        Assert.Equal("ByVal", parameters[1].PassingMode);
        Assert.Equal("c", parameters[2].Name);
        Assert.True(parameters[2].IsOptional);
    }

    [Fact]
    public void ParseModule_NoParameters_ReturnsEmptyParametersList()
    {
        var source = "Public Sub DoWork()\r\nEnd Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        Assert.Empty(result.Procedures.Single().Parameters);
    }

    [Fact]
    public void ParseModule_ContinuedParameterList_ParsesAllParametersWithCorrectPhysicalLines()
    {
        var source = "Public Sub DoWork(a As Long, _\r\n" +
                      "    b As String)\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var proc = result.Procedures.Single();
        Assert.Equal(2, proc.Parameters.Count);
        Assert.Equal("a", proc.Parameters[0].Name);
        Assert.Equal("b", proc.Parameters[1].Name);
        Assert.Equal(1, proc.StartLine);
        Assert.Equal(3, proc.EndLine);
    }

    [Fact]
    public void ParseModule_SingleDimDeclaration_ProducesOneLocal()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    Dim total As Long\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var local = Assert.Single(result.Procedures.Single().Locals);
        Assert.Equal("total", local.Name);
        Assert.Equal("Long", local.DeclaredType);
        Assert.Equal(SymbolKind.Local, local.Kind);
    }

    [Fact]
    public void ParseModule_MultiVariableDimLine_ProducesOneLocalPerVariable()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    Dim a As Long, b As String\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var locals = result.Procedures.Single().Locals;
        Assert.Equal(2, locals.Count);
        Assert.Equal("a", locals[0].Name);
        Assert.Equal("Long", locals[0].DeclaredType);
        Assert.Equal("b", locals[1].Name);
        Assert.Equal("String", locals[1].DeclaredType);
    }

    [Fact]
    public void ParseModule_StaticDeclaration_ProducesLocalKindLocal()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    Static counter As Long\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var local = Assert.Single(result.Procedures.Single().Locals);
        Assert.Equal("counter", local.Name);
        Assert.Equal(SymbolKind.Local, local.Kind);
    }

    [Fact]
    public void ParseModule_ConstDeclaration_ProducesKindConst()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    Const Max As Long = 100\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var local = Assert.Single(result.Procedures.Single().Locals);
        Assert.Equal("Max", local.Name);
        Assert.Equal(SymbolKind.Const, local.Kind);
        Assert.Equal("Long", local.DeclaredType);
    }

    [Fact]
    public void ParseModule_ArrayDeclaration_SetsIsArrayTrue()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    Dim values(10) As Long\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var local = Assert.Single(result.Procedures.Single().Locals);
        Assert.Equal("values", local.Name);
        Assert.True(local.IsArray);
        Assert.Equal("Long", local.DeclaredType);
    }

    [Fact]
    public void ParseModule_DynamicArrayDeclaration_SetsIsArrayTrue()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    Dim values() As Long\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var local = Assert.Single(result.Procedures.Single().Locals);
        Assert.True(local.IsArray);
    }

    [Fact]
    public void ParseModule_UntypedDim_DefaultsToVariant()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    Dim anything\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var local = Assert.Single(result.Procedures.Single().Locals);
        Assert.Equal("Variant", local.DeclaredType);
    }

    [Fact]
    public void ParseModule_NonDeclarationLinesInsideBody_AreSkippedNotErrored()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    Dim x As Long\r\n" +
                      "    x = 5\r\n" +
                      "    If x > 0 Then\r\n" +
                      "        x = x + 1\r\n" +
                      "    End If\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var local = Assert.Single(result.Procedures.Single().Locals);
        Assert.Equal("x", local.Name);
    }

    [Fact]
    public void ParseModule_NoLocalsInBody_ReturnsEmptyLocalsList()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    DoSomethingElse\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        Assert.Empty(result.Procedures.Single().Locals);
    }

    [Fact]
    public void ParseModule_ModuleLevelDimDeclaration_ProducesModuleVariable()
    {
        var source = "Option Explicit\r\n" +
                      "Dim counter As Long\r\n" +
                      "\r\n" +
                      "Public Sub DoWork()\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var moduleVar = Assert.Single(result.ModuleVariables);
        Assert.Equal("counter", moduleVar.Name);
        Assert.Equal(SymbolKind.ModuleVariable, moduleVar.Kind);
    }

    [Fact]
    public void ParseModule_ModuleLevelPublicDeclaration_NoDimKeywordNeeded()
    {
        var source = "Public sharedTotal As Long\r\n" +
                      "\r\n" +
                      "Public Sub DoWork()\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var moduleVar = Assert.Single(result.ModuleVariables);
        Assert.Equal("sharedTotal", moduleVar.Name);
        Assert.Equal(SymbolKind.ModuleVariable, moduleVar.Kind);
    }

    [Fact]
    public void ParseModule_ModuleLevelPrivateConst_ProducesKindConst()
    {
        var source = "Private Const MaxRetries As Long = 3\r\n" +
                      "\r\n" +
                      "Public Sub DoWork()\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var moduleVar = Assert.Single(result.ModuleVariables);
        Assert.Equal("MaxRetries", moduleVar.Name);
        Assert.Equal(SymbolKind.Const, moduleVar.Kind);
    }

    [Fact]
    public void ParseModule_ModuleLevelBareConst_DefaultsPrivateButStillKindConst()
    {
        var source = "Const Pi As Double = 3.14159\r\n" +
                      "\r\n" +
                      "Public Sub DoWork()\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        var moduleVar = Assert.Single(result.ModuleVariables);
        Assert.Equal("Pi", moduleVar.Name);
        Assert.Equal(SymbolKind.Const, moduleVar.Kind);
    }

    [Fact]
    public void ParseModule_MultipleModuleVariables_AllCollected()
    {
        var source = "Dim a As Long\r\n" +
                      "Private b As String\r\n" +
                      "Public c As Boolean\r\n" +
                      "\r\n" +
                      "Public Sub DoWork()\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        Assert.Equal(3, result.ModuleVariables.Count);
        Assert.Contains(result.ModuleVariables, s => s.Name == "a");
        Assert.Contains(result.ModuleVariables, s => s.Name == "b");
        Assert.Contains(result.ModuleVariables, s => s.Name == "c");
    }

    [Fact]
    public void ParseModule_DeclarationsInsideProcedureBody_AreNotCountedAsModuleVariables()
    {
        var source = "Public Sub DoWork()\r\n" +
                      "    Dim local1 As Long\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        Assert.Empty(result.ModuleVariables);
        Assert.Single(result.Procedures.Single().Locals);
    }

    [Fact]
    public void ParseModule_NoModuleLevelVariables_ReturnsEmptyList()
    {
        var source = "Option Explicit\r\n" +
                      "\r\n" +
                      "Public Sub DoWork()\r\n" +
                      "End Sub\r\n";

        var result = VbaParser.ParseModule(source, "modWork");

        Assert.Empty(result.ModuleVariables);
    }
}
