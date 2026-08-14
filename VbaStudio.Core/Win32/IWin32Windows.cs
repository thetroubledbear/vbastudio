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
}
