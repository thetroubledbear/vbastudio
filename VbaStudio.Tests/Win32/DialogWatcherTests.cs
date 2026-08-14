using System;
using System.Collections.Generic;
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
    public void PollOnce_MatchingCaptionButNotADialogClass_Ignored()
    {
        // Excel's own frame window (XLMAIN) carries the caption "Microsoft Excel" in some
        // states; only the standard dialog class #32770 may be enumerated and clicked.
        var fake = new FakeWin32Windows();
        fake.AddTopLevelWindow(new FakeWindow
        {
            Handle = (IntPtr)1, Caption = "Microsoft Excel",
            ClassName = "XLMAIN", ProcessId = TargetPid
        });
        fake.AddChildWindow((IntPtr)1, new FakeWindow { Handle = (IntPtr)10, Caption = "Some frame text", ClassName = "Static", ProcessId = TargetPid });
        fake.AddChildWindow((IntPtr)1, new FakeWindow { Handle = (IntPtr)11, Caption = "OK", ClassName = "Button", ProcessId = TargetPid });

        var watcher = new DialogWatcher(fake, TargetPid);
        watcher.PollOnce();

        Assert.Null(watcher.Captured);
        Assert.Empty(fake.ClickedHandles);
    }

    [Fact]
    public void PollOnce_HelpButtonBeforeOkInZOrder_ClicksOk()
    {
        var fake = new FakeWin32Windows();
        fake.AddTopLevelWindow(new FakeWindow
        {
            Handle = (IntPtr)1, Caption = "Microsoft Visual Basic for Applications",
            ClassName = "#32770", ProcessId = TargetPid
        });
        fake.AddChildWindow((IntPtr)1, new FakeWindow { Handle = (IntPtr)10, Caption = "Compile error:", ClassName = "Static", ProcessId = TargetPid });
        // Help is enumerated first - EnumChildWindows returns Z-order, not OK-first order.
        fake.AddChildWindow((IntPtr)1, new FakeWindow { Handle = (IntPtr)20, Caption = "Help", ClassName = "Button", ProcessId = TargetPid });
        fake.AddChildWindow((IntPtr)1, new FakeWindow { Handle = (IntPtr)21, Caption = "OK", ClassName = "Button", ProcessId = TargetPid });

        var watcher = new DialogWatcher(fake, TargetPid);
        watcher.PollOnce();

        Assert.NotNull(watcher.Captured);
        Assert.Equal(new[] { (IntPtr)21 }, fake.ClickedHandles);
    }

    [Fact]
    public void PollOnce_OnlyNonOkButton_FallsBackToFirstButton()
    {
        var fake = new FakeWin32Windows();
        fake.AddTopLevelWindow(new FakeWindow
        {
            Handle = (IntPtr)1, Caption = "Microsoft Excel",
            ClassName = "#32770", ProcessId = TargetPid
        });
        fake.AddChildWindow((IntPtr)1, new FakeWindow { Handle = (IntPtr)10, Caption = "Continue?", ClassName = "Static", ProcessId = TargetPid });
        fake.AddChildWindow((IntPtr)1, new FakeWindow { Handle = (IntPtr)20, Caption = "Continue", ClassName = "Button", ProcessId = TargetPid });

        var watcher = new DialogWatcher(fake, TargetPid);
        watcher.PollOnce();

        Assert.Equal(new[] { (IntPtr)20 }, fake.ClickedHandles);
    }

    [Fact]
    public void PollOnce_MatchedDialog_IsLogged()
    {
        var fake = new FakeWin32Windows();
        fake.AddTopLevelWindow(new FakeWindow
        {
            Handle = (IntPtr)1, Caption = "Microsoft Visual Basic for Applications",
            ClassName = "#32770", ProcessId = TargetPid
        });
        fake.AddChildWindow((IntPtr)1, new FakeWindow { Handle = (IntPtr)10, Caption = "Syntax error", ClassName = "Static", ProcessId = TargetPid });

        var logged = new List<string>();
        var watcher = new DialogWatcher(fake, TargetPid, logged.Add);
        watcher.PollOnce();

        var line = Assert.Single(logged);
        Assert.Contains("Microsoft Visual Basic for Applications", line);
        Assert.Contains("Syntax error", line);
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
