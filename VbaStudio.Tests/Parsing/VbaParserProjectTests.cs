// VbaStudio.Tests/Parsing/VbaParserProjectTests.cs
using System.IO.Abstractions.TestingHelpers;
using System.Linq;
using VbaStudio.Core.Parsing;
using Xunit;

namespace VbaStudio.Tests.Parsing;

public class VbaParserProjectTests
{
    [Fact]
    public void ParseProject_ScansAllFourSubfolders()
    {
        var fs = new MockFileSystem();
        fs.AddFile("src/Modules/modWork.bas", new MockFileData(
            "Public Sub DoWork()\r\nEnd Sub\r\n"));
        fs.AddFile("src/Classes/clsThing.cls", new MockFileData(
            "Public Sub DoThing()\r\nEnd Sub\r\n"));
        fs.AddFile("src/Forms/frmMain.frm", new MockFileData(
            "Public Sub Init()\r\nEnd Sub\r\n"));
        fs.AddFile("src/Sheets/Sheet1.cls", new MockFileData(
            "Public Sub OnLoad()\r\nEnd Sub\r\n"));

        var result = VbaParser.ParseProject(fs, "src");

        Assert.Equal(4, result.Count);
        Assert.Contains(result, m => m.ModuleName == "modWork");
        Assert.Contains(result, m => m.ModuleName == "clsThing");
        Assert.Contains(result, m => m.ModuleName == "frmMain");
        Assert.Contains(result, m => m.ModuleName == "Sheet1");
    }

    [Fact]
    public void ParseProject_NoFilenameFilter_ParsesEveryFileRegardlessOfName()
    {
        var fs = new MockFileSystem();
        fs.AddFile("src/Modules/modAnything.bas", new MockFileData(
            "Public Sub Helper()\r\nEnd Sub\r\n"));

        var result = VbaParser.ParseProject(fs, "src");

        var module = Assert.Single(result);
        Assert.Equal("modAnything", module.ModuleName);
        Assert.Single(module.Procedures);
    }

    [Fact]
    public void ParseProject_MissingSubfolders_ReturnsEmptyWithoutThrowing()
    {
        var fs = new MockFileSystem();

        var result = VbaParser.ParseProject(fs, "src");

        Assert.Empty(result);
    }

    [Fact]
    public void ParseProject_ModuleSymbolsMatchDirectParseModuleCall()
    {
        var fs = new MockFileSystem();
        fs.AddFile("src/Modules/modWork.bas", new MockFileData(
            "Public Sub DoWork(a As Long)\r\n    Dim total As Long\r\nEnd Sub\r\n"));

        var result = VbaParser.ParseProject(fs, "src");

        var viaProject = result.Single();
        var viaDirect = VbaParser.ParseModule(
            "Public Sub DoWork(a As Long)\r\n    Dim total As Long\r\nEnd Sub\r\n", "modWork");

        Assert.Equal(viaDirect.Procedures.Single().Name, viaProject.Procedures.Single().Name);
        Assert.Equal(viaDirect.Procedures.Single().Parameters.Count, viaProject.Procedures.Single().Parameters.Count);
        Assert.Equal(viaDirect.Procedures.Single().Locals.Count, viaProject.Procedures.Single().Locals.Count);
    }
}
