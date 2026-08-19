using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace VbaStudio.Core.Win32;

public sealed class Win32Windows : IWin32Windows
{
    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr hwndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    private static extern bool IsWindow(IntPtr hWnd);

    // PostMessage, not SendMessage: SendMessage blocks until the target window's thread
    // pumps the message, so an unresponsive Excel would hang the watcher thread here (and
    // then hang the runner thread inside DialogWatcher.Stop()). PostMessage queues and returns.
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    // keybd_event, not SendInput: this process injects exactly one fixed key combo into whatever
    // window currently has focus - SendInput's richer (and more boilerplate-heavy) INPUT-array API
    // buys nothing here that keybd_event doesn't already do in four calls.
    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    private const uint BM_CLICK = 0x00F5;
    private const byte VK_CONTROL = 0x11;
    private const byte VK_CANCEL = 0x03; // Ctrl+Break's virtual-key code - distinct from VK_PAUSE (0x13).
    private const uint KEYEVENTF_KEYUP = 0x0002;

    public IReadOnlyList<IntPtr> EnumerateTopLevelWindows()
    {
        var result = new List<IntPtr>();
        EnumWindows((hwnd, _) => { result.Add(hwnd); return true; }, IntPtr.Zero);
        return result;
    }

    public IReadOnlyList<IntPtr> EnumerateChildWindows(IntPtr parent)
    {
        var result = new List<IntPtr>();
        EnumChildWindows(parent, (hwnd, _) => { result.Add(hwnd); return true; }, IntPtr.Zero);
        return result;
    }

    public string GetWindowText(IntPtr hwnd)
    {
        var buffer = new StringBuilder(512);
        GetWindowText(hwnd, buffer, buffer.Capacity);
        return buffer.ToString();
    }

    public string GetClassName(IntPtr hwnd)
    {
        var buffer = new StringBuilder(256);
        GetClassName(hwnd, buffer, buffer.Capacity);
        return buffer.ToString();
    }

    public int GetWindowProcessId(IntPtr hwnd)
    {
        GetWindowThreadProcessId(hwnd, out uint processId);
        return (int)processId;
    }

    public void Click(IntPtr hwndButton)
    {
        PostMessage(hwndButton, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
    }

    public bool WaitForWindowClosed(IntPtr hwnd, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (IsWindow(hwnd))
        {
            if (DateTime.UtcNow >= deadline)
            {
                return false;
            }

            Thread.Sleep(20);
        }

        return true;
    }

    public void SendCtrlBreak(IntPtr targetWindow)
    {
        SetForegroundWindow(targetWindow);
        keybd_event(VK_CONTROL, 0, 0, UIntPtr.Zero);
        keybd_event(VK_CANCEL, 0, 0, UIntPtr.Zero);
        keybd_event(VK_CANCEL, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }
}
