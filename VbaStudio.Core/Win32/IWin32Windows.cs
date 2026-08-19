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

    /// <summary>
    /// Simulates a physical Ctrl+Break key press, after bringing <paramref name="targetWindow"/> to
    /// the foreground - VBA's own Ctrl+Break interrupt detection (Application.EnableCancelKey) only
    /// fires while Excel is the active application, mirroring the manual "click into Excel, press
    /// Esc" experience. Used to force a stop request through to a macro that Application.Run is
    /// currently blocked inside, with nothing else able to interrupt it - see DapRequestReader.
    /// </summary>
    void SendCtrlBreak(IntPtr targetWindow);
}
