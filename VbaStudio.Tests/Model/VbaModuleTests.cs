using VbaStudio.Core.Model;
using Xunit;

namespace VbaStudio.Tests.Model;

public class VbaModuleTests
{
    [Theory]
    [InlineData(ModuleKind.Standard, ".bas")]
    [InlineData(ModuleKind.Class, ".cls")]
    [InlineData(ModuleKind.UserForm, ".frm")]
    [InlineData(ModuleKind.Document, ".cls")]
    public void FileExtension_MatchesModuleKind(ModuleKind kind, string expected)
    {
        Assert.Equal(expected, kind.FileExtension());
    }

    [Fact]
    public void FileName_CombinesNameAndExtension()
    {
        var module = new VbaModule("modCalc", ModuleKind.Standard, "Sub Foo()\r\nEnd Sub", ".bas");
        Assert.Equal("modCalc.bas", module.FileName);
    }

    [Fact]
    public void IsDocumentModule_TrueOnlyForDocumentKind()
    {
        var doc = new VbaModule("Sheet1", ModuleKind.Document, "", ".cls");
        var cls = new VbaModule("clsFoo", ModuleKind.Class, "", ".cls");
        Assert.True(doc.IsDocumentModule);
        Assert.False(cls.IsDocumentModule);
    }
}
