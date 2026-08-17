using System;
using System.Collections.Generic;

namespace VbaStudio.Core.Win32;

public interface IWin32Windows
{
    IReadOnlyList<IntPtr> EnumerateTopLevelWindows();
    IReadOnlyList<IntPtr> EnumerateChildWindows(IntPtr parent);
    string GetWindowText(IntPtr hwnd);
    string GetClassName(IntPtr hwnd);
    int GetWindowProcessId(IntPtr hwnd);
    void Click(IntPtr hwndButton);

    /// <summary>
    /// Blocks (on the calling thread only - never a COM thread, per DialogWatcher's threading
    /// contract) until <paramref name="hwnd"/> is no longer a valid window, or the timeout
    /// elapses. Click() posts asynchronously and returns immediately; without this, the next
    /// operation's watcher can catch the same dialog still closing and misreport it as fresh.
    /// </summary>
    bool WaitForWindowClosed(IntPtr hwnd, TimeSpan timeout);
}
