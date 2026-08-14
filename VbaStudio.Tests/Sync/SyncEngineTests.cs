using System;
using System.IO.Abstractions.TestingHelpers;
using System.Linq;
using System.Text;
using VbaStudio.Core.Excel;
using VbaStudio.Core.Model;
using VbaStudio.Core.Sync;
using VbaStudio.Tests.Fakes;
using Xunit;

namespace VbaStudio.Tests.Sync;

public class SyncEngineTests
{
    [Fact]
    public void PullThenPush_RoundTrip_ContentUnchanged()
    {
        var fake = new FakeVbaProjectAccess();
        fake.Add(new VbaModule("modCalc", ModuleKind.Standard, "Attribute VB_Name = \"modCalc\"\r\nSub Foo()\r\nEnd Sub", ".bas"));
        var fs = new MockFileSystem();
        var sync = new SyncEngine(fake, fs, "src");

        sync.Pull();
        var pulled = fs.File.ReadAllText(fs.Path.Combine("src", "Modules", "modCalc.bas"), Encoding.GetEncoding(1252));
        Assert.Equal(fake.ReadAll().Single().Code, pulled);

        sync.Push();
        sync.Pull();
        var pulledAgain = fs.File.ReadAllText(fs.Path.Combine("src", "Modules", "modCalc.bas"), Encoding.GetEncoding(1252));

        Assert.Equal(pulled, pulledAgain);
    }

    [Fact]
    public void Pull_DocumentModule_UsesUtf8Encoding()
    {
        var fake = new FakeVbaProjectAccess();
        fake.Add(new VbaModule("Sheet1", ModuleKind.Document, "Option Explicit\r\n' Café", ".cls"));
        var fs = new MockFileSystem();
        var sync = new SyncEngine(fake, fs, "src");

        sync.Pull();

        var bytes = fs.File.ReadAllBytes(fs.Path.Combine("src", "Sheets", "Sheet1.cls"));
        var text = Encoding.UTF8.GetString(bytes);
        Assert.Contains("Café", text);
    }

    [Fact]
    public void Push_MacroRunning_Throws()
    {
        var fake = new FakeVbaProjectAccess { IsMacroRunning = true };
        var fs = new MockFileSystem();
        fs.AddDirectory("src");
        var sync = new SyncEngine(fake, fs, "src");

        Assert.Throws<InvalidOperationException>(() => sync.Push());
    }

    [Fact]
    public void Push_SrcDirMissing_CreatesDirectoryAndReturnsWithoutThrowing()
    {
        var fake = new FakeVbaProjectAccess();
        var fs = new MockFileSystem();
        var sync = new SyncEngine(fake, fs, "src");

        var exception = Record.Exception(() => sync.Push());

        Assert.Null(exception);
        Assert.True(fs.Directory.Exists("src"));
        Assert.Equal(0, fake.WriteCallCount);
    }

    [Fact]
    public void Push_UnchangedContent_DoesNotCallWrite()
    {
        var fake = new FakeVbaProjectAccess();
        fake.Add(new VbaModule("modCalc", ModuleKind.Standard, "Sub Foo()\r\nEnd Sub", ".bas"));
        var fs = new MockFileSystem();
        var sync = new SyncEngine(fake, fs, "src");
        sync.Pull();

        sync.Push();

        Assert.Equal(0, fake.WriteCallCount);
    }

    [Fact]
    public void Push_ChangedContent_CallsWriteWithNewCode()
    {
        var fake = new FakeVbaProjectAccess();
        fake.Add(new VbaModule("modCalc", ModuleKind.Standard, "Sub Foo()\r\nEnd Sub", ".bas"));
        var fs = new MockFileSystem();
        var sync = new SyncEngine(fake, fs, "src");
        sync.Pull();
        fs.File.WriteAllText(fs.Path.Combine("src", "Modules", "modCalc.bas"), "Sub Foo()\r\n    ' changed\r\nEnd Sub", Encoding.GetEncoding(1252));

        sync.Push();

        Assert.Equal(1, fake.WriteCallCount);
        Assert.Contains("changed", fake.ReadAll().Single().Code);
    }

    [Fact]
    public void Push_IgnoresUnrecognizedExtensions()
    {
        var fake = new FakeVbaProjectAccess();
        var fs = new MockFileSystem();
        fs.AddFile("src/Forms/UserFormMain.frx", new MockFileData(new byte[] { 0x01, 0x02 }));
        fs.AddFile("src/.gitattributes", new MockFileData("*.frx binary"));
        var sync = new SyncEngine(fake, fs, "src");

        var exception = Record.Exception(() => sync.Push());

        Assert.Null(exception);
        Assert.Equal(0, fake.WriteCallCount);
    }
}
