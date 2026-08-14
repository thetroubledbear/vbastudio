using System;
using VbaStudio.Core.Win32;
using VbaStudio.Tests.Fakes;
using Xunit;

namespace VbaStudio.Tests.Win32;

public class DialogWatcherTests
{
    private const int TargetPid = 4242;

    [Fact]
    public void PollOnce_CompileErrorDialog_CapturesConcatenatedBodyAndClicksButton()
    {
        var fake = new FakeWin32Windows();
        fake.AddTopLevelWindow(new FakeWindow
        {
            Handle = (IntPtr)1, Caption = "Microsoft Visual Basic for Applications",
            ClassName = "#32770", ProcessId = TargetPid
        });
        fake.AddChildWindow((IntPtr)1, new FakeWindow { Handle = (IntPtr)10, Caption = "Compile error:", ClassName = "Static", ProcessId = TargetPid });
        fake.AddChildWindow((IntPtr)1, new FakeWindow { Handle = (IntPtr)11, Caption = "Syntax error", ClassName = "Static", ProcessId = TargetPid });
        fake.AddChildWindow((IntPtr)1, new FakeWindow { Handle = (IntPtr)12, Caption = "OK", ClassName = "Button", ProcessId = TargetPid });

        var watcher = new DialogWatcher(fake, TargetPid);
        watcher.PollOnce();

        Assert.NotNull(watcher.Captured);
        Assert.Equal("Microsoft Visual Basic for Applications", watcher.Captured!.Caption);
        Assert.Equal("Compile error: Syntax error", watcher.Captured!.Body);
        Assert.Equal(new[] { (IntPtr)12 }, fake.ClickedHandles);
    }

    [Fact]
    public void PollOnce_DifferentProcessId_Ignored()
    {
        var fake = new FakeWin32Windows();
        fake.AddTopLevelWindow(new FakeWindow
        {
            Handle = (IntPtr)1, Caption = "Microsoft Visual Basic for Applications",
            ClassName = "#32770", ProcessId = 9999
        });

        var watcher = new DialogWatcher(fake, TargetPid);
        watcher.PollOnce();

        Assert.Null(watcher.Captured);
        Assert.Empty(fake.ClickedHandles);
    }

    [Fact]
    public void PollOnce_UnrecognizedCaption_Ignored()
    {
        var fake = new FakeWin32Windows();
        fake.AddTopLevelWindow(new FakeWindow
        {
            Handle = (IntPtr)1, Caption = "Some Other App",
            ClassName = "#32770", ProcessId = TargetPid
        });

        var watcher = new DialogWatcher(fake, TargetPid);
        watcher.PollOnce();

        Assert.Null(watcher.Captured);
        Assert.Empty(fake.ClickedHandles);
    }

    [Fact]
    public void PollOnce_KnownCaptionNoButtonChild_CapturesButDoesNotClick()
    {
        var fake = new FakeWin32Windows();
        fake.AddTopLevelWindow(new FakeWindow
        {
            Handle = (IntPtr)1, Caption = "Microsoft Excel",
            ClassName = "#32770", ProcessId = TargetPid
        });
        fake.AddChildWindow((IntPtr)1, new FakeWindow { Handle = (IntPtr)10, Caption = "This will reset your project.", ClassName = "Static", ProcessId = TargetPid });

        var watcher = new DialogWatcher(fake, TargetPid);
        watcher.PollOnce();

        Assert.NotNull(watcher.Captured);
        Assert.Equal("This will reset your project.", watcher.Captured!.Body);
        Assert.Empty(fake.ClickedHandles);
    }

    [Fact]
    public void PollOnce_NoMatchingWindow_CapturedStaysNull()
    {
        var fake = new FakeWin32Windows();

        var watcher = new DialogWatcher(fake, TargetPid);
        watcher.PollOnce();

        Assert.Null(watcher.Captured);
    }
}
