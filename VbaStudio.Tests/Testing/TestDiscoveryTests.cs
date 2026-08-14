using System.IO.Abstractions.TestingHelpers;
using VbaStudio.Core.Testing;
using Xunit;

namespace VbaStudio.Tests.Testing;

public class TestDiscoveryTests
{
    [Fact]
    public void DiscoverTests_FindsTestProceduresInMatchingClsFile()
    {
        var fs = new MockFileSystem();
        // Only src/Modules is scanned (Application.Run cannot invoke class-module procedures), but
        // the file-name pattern still accepts a .cls extension - that half of the regex is covered here.
        fs.AddFile("src/Modules/modMathTests.cls", new MockFileData(
            "Option Explicit\r\n" +
            "\r\n" +
            "Public Sub Test_AddsTwoNumbers()\r\n" +
            "    Dim asserter As New clsAssert\r\n" +
            "    asserter.AreEqual 4, 2 + 2\r\n" +
            "End Sub\r\n" +
            "\r\n" +
            "Public Sub Test_SubtractsTwoNumbers()\r\n" +
            "    Dim asserter As New clsAssert\r\n" +
            "    asserter.AreEqual 0, 2 - 2\r\n" +
            "End Sub\r\n"));

        var result = TestDiscovery.DiscoverTests(fs, "src");

        Assert.Equal(2, result.Count);
        Assert.Contains(result, t => t.ModuleName == "modMathTests" && t.ProcedureName == "Test_AddsTwoNumbers");
        Assert.Contains(result, t => t.ModuleName == "modMathTests" && t.ProcedureName == "Test_SubtractsTwoNumbers");
    }

    [Fact]
    public void DiscoverTests_FindsTestProceduresInMatchingBasFile()
    {
        var fs = new MockFileSystem();
        fs.AddFile("src/Modules/modStringTests.bas", new MockFileData(
            "Attribute VB_Name = \"modStringTests\"\r\n" +
            "Public Sub Test_ConcatenatesStrings()\r\n" +
            "    Dim asserter As New clsAssert\r\n" +
            "    asserter.AreEqual \"ab\", \"a\" & \"b\"\r\n" +
            "End Sub\r\n"));

        var result = TestDiscovery.DiscoverTests(fs, "src");

        var testCase = Assert.Single(result);
        Assert.Equal("modStringTests", testCase.ModuleName);
        Assert.Equal("Test_ConcatenatesStrings", testCase.ProcedureName);
        Assert.Equal("modStringTests.Test_ConcatenatesStrings", testCase.QualifiedName);
    }

    [Fact]
    public void DiscoverTests_IgnoresFilesNotMatchingTestsNamingConvention()
    {
        var fs = new MockFileSystem();
        fs.AddFile("src/Modules/modMath.bas", new MockFileData(
            "Public Sub Test_ThisLooksLikeATestButIsNot()\r\n" +
            "End Sub\r\n"));

        var result = TestDiscovery.DiscoverTests(fs, "src");

        Assert.Empty(result);
    }

    [Fact]
    public void DiscoverTests_IgnoresNonTestProceduresInATestsFile()
    {
        var fs = new MockFileSystem();
        fs.AddFile("src/Modules/modMathTests.bas", new MockFileData(
            "Public Sub Test_RealTest()\r\n" +
            "    Dim asserter As New clsAssert\r\n" +
            "    asserter.IsTrue True\r\n" +
            "End Sub\r\n" +
            "\r\n" +
            "Private Sub HelperNotATest()\r\n" +
            "End Sub\r\n" +
            "\r\n" +
            "Public Sub NotPrefixedWithTest_Underscore()\r\n" +
            "End Sub\r\n"));

        var result = TestDiscovery.DiscoverTests(fs, "src");

        var testCase = Assert.Single(result);
        Assert.Equal("Test_RealTest", testCase.ProcedureName);
    }

    [Fact]
    public void DiscoverTests_IgnoresTestProcedureWithParameters()
    {
        var fs = new MockFileSystem();
        fs.AddFile("src/Modules/modMathTests.bas", new MockFileData(
            "Public Sub Test_WithParams(x As Integer)\r\n" +
            "End Sub\r\n" +
            "\r\n" +
            "Public Sub Test_NoParams()\r\n" +
            "End Sub\r\n"));

        var result = TestDiscovery.DiscoverTests(fs, "src");

        var testCase = Assert.Single(result);
        Assert.Equal("Test_NoParams", testCase.ProcedureName);
    }

    [Fact]
    public void DiscoverTests_IgnoresClassModulesBecauseApplicationRunCannotInvokeThem()
    {
        var fs = new MockFileSystem();
        fs.AddFile("src/Classes/clsMathTests.cls", new MockFileData(
            "Public Sub Test_InAClassModule()\r\n" +
            "End Sub\r\n"));

        var result = TestDiscovery.DiscoverTests(fs, "src");

        Assert.Empty(result);
    }

    [Fact]
    public void DiscoverTests_MissingSrcSubfolders_ReturnsEmptyWithoutThrowing()
    {
        var fs = new MockFileSystem();

        var result = TestDiscovery.DiscoverTests(fs, "src");

        Assert.Empty(result);
    }
}
