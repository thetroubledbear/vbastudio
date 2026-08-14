using System;
using System.Collections.Generic;
using System.Linq;
using VbaStudio.Core.Win32;

namespace VbaStudio.Tests.Fakes;

public sealed class FakeWindow
{
    public IntPtr Handle { get; init; }
    public string Caption { get; init; } = "";
    public string ClassName { get; init; } = "";
    public int ProcessId { get; init; }
    public List<IntPtr> Children { get; } = new();
}

public sealed class FakeWin32Windows : IWin32Windows
{
    private readonly Dictionary<IntPtr, FakeWindow> _windows = new();
    private readonly List<IntPtr> _topLevel = new();
    private readonly List<IntPtr> _clickedHandles = new();
    private readonly HashSet<IntPtr> _neverCloses = new();

    public IReadOnlyList<IntPtr> ClickedHandles => _clickedHandles;

    /// <summary>Makes WaitForWindowClosed report the window still open until the timeout elapses.</summary>
    public void MarkNeverCloses(IntPtr hwnd) => _neverCloses.Add(hwnd);

    public void AddTopLevelWindow(FakeWindow window)
    {
        _windows[window.Handle] = window;
        _topLevel.Add(window.Handle);
    }

    public void AddChildWindow(IntPtr parentHandle, FakeWindow child)
    {
        _windows[child.Handle] = child;
        _windows[parentHandle].Children.Add(child.Handle);
    }

    public IReadOnlyList<IntPtr> EnumerateTopLevelWindows() => _topLevel;

    public IReadOnlyList<IntPtr> EnumerateChildWindows(IntPtr parent) =>
        _windows.TryGetValue(parent, out var window) ? window.Children : Array.Empty<IntPtr>();

    public string GetWindowText(IntPtr hwnd) => _windows[hwnd].Caption;

    public string GetClassName(IntPtr hwnd) => _windows[hwnd].ClassName;

    public int GetWindowProcessId(IntPtr hwnd) => _windows[hwnd].ProcessId;

    public void Click(IntPtr hwndButton) => _clickedHandles.Add(hwndButton);

    public bool WaitForWindowClosed(IntPtr hwnd, TimeSpan timeout) => !_neverCloses.Contains(hwnd);
}
