using System;
using System.Linq;
using VbaStudio.Core.Win32;
using Xunit;

namespace VbaStudio.Tests.Fakes;

public class FakeWin32WindowsTests
{
    [Fact]
    public void EnumerateTopLevelWindows_ReturnsAddedHandles()
    {
        var fake = new FakeWin32Windows();
        fake.AddTopLevelWindow(new FakeWindow { Handle = (IntPtr)1, Caption = "A", ClassName = "#32770", ProcessId = 100 });
        fake.AddTopLevelWindow(new FakeWindow { Handle = (IntPtr)2, Caption = "B", ClassName = "#32770", ProcessId = 100 });

        var result = fake.EnumerateTopLevelWindows();

        Assert.Equal(new[] { (IntPtr)1, (IntPtr)2 }, result);
    }

    [Fact]
    public void EnumerateChildWindows_ReturnsChildrenOfSpecificParent()
    {
        var fake = new FakeWin32Windows();
        fake.AddTopLevelWindow(new FakeWindow { Handle = (IntPtr)1, Caption = "Dialog", ClassName = "#32770", ProcessId = 100 });
        fake.AddChildWindow((IntPtr)1, new FakeWindow { Handle = (IntPtr)10, Caption = "OK", ClassName = "Button", ProcessId = 100 });
        fake.AddChildWindow((IntPtr)1, new FakeWindow { Handle = (IntPtr)11, Caption = "Compile error", ClassName = "Static", ProcessId = 100 });

        var children = fake.EnumerateChildWindows((IntPtr)1);

        Assert.Equal(new[] { (IntPtr)10, (IntPtr)11 }, children);
    }

    [Fact]
    public void GetWindowText_GetClassName_GetWindowProcessId_ReturnRegisteredValues()
    {
        var fake = new FakeWin32Windows();
        fake.AddTopLevelWindow(new FakeWindow { Handle = (IntPtr)1, Caption = "Microsoft Excel", ClassName = "#32770", ProcessId = 4242 });

        Assert.Equal("Microsoft Excel", fake.GetWindowText((IntPtr)1));
        Assert.Equal("#32770", fake.GetClassName((IntPtr)1));
        Assert.Equal(4242, fake.GetWindowProcessId((IntPtr)1));
    }

    [Fact]
    public void Click_RecordsHandleInClickedHandles()
    {
        var fake = new FakeWin32Windows();

        fake.Click((IntPtr)10);
        fake.Click((IntPtr)11);

        Assert.Equal(new[] { (IntPtr)10, (IntPtr)11 }, fake.ClickedHandles);
    }
}
