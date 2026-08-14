using System.IO;
using System.IO.Abstractions.TestingHelpers;
using VbaStudio.Core.Shadow;
using Xunit;

namespace VbaStudio.Tests.Shadow;

public class ShadowWorkbookTests
{
    [Fact]
    public void CreateFromClosed_CopiesSourceContentToShadowPath()
    {
        var fs = new MockFileSystem();
        fs.AddFile("C:/work/Reporting.xlsm", new MockFileData("fake workbook bytes"));

        ShadowWorkbook.CreateFromClosed(fs, "C:/work/Reporting.xlsm", "C:/work/build/shadow.xlsm");

        Assert.True(fs.FileExists("C:/work/build/shadow.xlsm"));
        Assert.Equal("fake workbook bytes", fs.File.ReadAllText("C:/work/build/shadow.xlsm"));
    }

    [Fact]
    public void CreateFromClosed_OverwritesExistingShadowFile()
    {
        var fs = new MockFileSystem();
        fs.AddFile("C:/work/Reporting.xlsm", new MockFileData("new content"));
        fs.AddFile("C:/work/build/shadow.xlsm", new MockFileData("stale content"));

        ShadowWorkbook.CreateFromClosed(fs, "C:/work/Reporting.xlsm", "C:/work/build/shadow.xlsm");

        Assert.Equal("new content", fs.File.ReadAllText("C:/work/build/shadow.xlsm"));
    }

    [Fact]
    public void CreateFromClosed_MissingSourcePath_Throws()
    {
        var fs = new MockFileSystem();

        Assert.Throws<FileNotFoundException>(() =>
            ShadowWorkbook.CreateFromClosed(fs, "C:/work/DoesNotExist.xlsm", "C:/work/build/shadow.xlsm"));
    }
}
