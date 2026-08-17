using System.Linq;
using VbaStudio.Core.Model;
using VbaStudio.Core.Tooling;
using Xunit;

namespace VbaStudio.Tests.Tooling;

public class ModuleListBuilderTests
{
    [Fact]
    public void Build_ModuleWithTwoProcedures_ListsBothInDeclarationOrder()
    {
        var source = "Public Sub First()\r\n" +
                      "End Sub\r\n" +
                      "\r\n" +
                      "Public Function Second() As Long\r\n" +
                      "    Second = 1\r\n" +
                      "End Function\r\n";
        var modules = new[] { new VbaModule("modWork", ModuleKind.Standard, source, ".bas") };

        var result = ModuleListBuilder.Build(@"C:\work\Reporting.xlsm", modules);

        Assert.Equal(@"C:\work\Reporting.xlsm", result.WorkbookPath);
        var listing = Assert.Single(result.Modules);
        Assert.Equal("modWork", listing.Name);
        Assert.Equal(new[] { "First", "Second" }, listing.Procedures);
    }

    [Fact]
    public void Build_ModuleWithNoProcedures_ReturnsEmptyProceduresList()
    {
        var source = "Option Explicit\r\n";
        var modules = new[] { new VbaModule("modEmpty", ModuleKind.Standard, source, ".bas") };

        var result = ModuleListBuilder.Build(@"C:\work\Reporting.xlsm", modules);

        var listing = Assert.Single(result.Modules);
        Assert.Empty(listing.Procedures);
    }

    [Fact]
    public void Build_MultipleModules_PreservesModuleOrder()
    {
        var modules = new[]
        {
            new VbaModule("modA", ModuleKind.Standard, "Public Sub A()\r\nEnd Sub\r\n", ".bas"),
            new VbaModule("modB", ModuleKind.Standard, "Public Sub B()\r\nEnd Sub\r\n", ".bas"),
        };

        var result = ModuleListBuilder.Build(@"C:\work\Reporting.xlsm", modules);

        Assert.Equal(new[] { "modA", "modB" }, result.Modules.Select(m => m.Name));
    }

    [Fact]
    public void Build_ClassModule_ExcludedEntirely()
    {
        var source = "Public Sub DoWork()\r\nEnd Sub\r\n";
        var modules = new[] { new VbaModule("clsThing", ModuleKind.Class, source, ".cls") };

        var result = ModuleListBuilder.Build(@"C:\work\Reporting.xlsm", modules);

        Assert.Empty(result.Modules);
    }

    [Fact]
    public void Build_ProcedureWithRequiredParameter_Excluded()
    {
        var source = "Public Sub NeedsArg(x As Long)\r\nEnd Sub\r\n" +
                      "Public Sub NoArgs()\r\nEnd Sub\r\n";
        var modules = new[] { new VbaModule("modWork", ModuleKind.Standard, source, ".bas") };

        var result = ModuleListBuilder.Build(@"C:\work\Reporting.xlsm", modules);

        var listing = Assert.Single(result.Modules);
        Assert.Equal(new[] { "NoArgs" }, listing.Procedures);
    }

    [Fact]
    public void Build_PropertyGet_Excluded()
    {
        var source = "Public Property Get Foo() As Long\r\n" +
                      "    Foo = 1\r\n" +
                      "End Property\r\n" +
                      "Public Sub NoArgs()\r\nEnd Sub\r\n";
        var modules = new[] { new VbaModule("modWork", ModuleKind.Standard, source, ".bas") };

        var result = ModuleListBuilder.Build(@"C:\work\Reporting.xlsm", modules);

        var listing = Assert.Single(result.Modules);
        Assert.Equal(new[] { "NoArgs" }, listing.Procedures);
    }

    [Fact]
    public void Build_PrivateSub_Excluded()
    {
        var source = "Private Sub Helper()\r\nEnd Sub\r\n" +
                      "Public Sub NoArgs()\r\nEnd Sub\r\n";
        var modules = new[] { new VbaModule("modWork", ModuleKind.Standard, source, ".bas") };

        var result = ModuleListBuilder.Build(@"C:\work\Reporting.xlsm", modules);

        var listing = Assert.Single(result.Modules);
        Assert.Equal(new[] { "NoArgs" }, listing.Procedures);
    }

    [Fact]
    public void Build_ProcedureWithOnlyOptionalParameters_StillIncluded()
    {
        var source = "Public Sub MaybeArg(Optional x As Long)\r\nEnd Sub\r\n";
        var modules = new[] { new VbaModule("modWork", ModuleKind.Standard, source, ".bas") };

        var result = ModuleListBuilder.Build(@"C:\work\Reporting.xlsm", modules);

        var listing = Assert.Single(result.Modules);
        Assert.Equal(new[] { "MaybeArg" }, listing.Procedures);
    }
}
