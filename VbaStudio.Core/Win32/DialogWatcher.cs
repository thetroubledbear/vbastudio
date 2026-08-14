using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace VbaStudio.Core.Win32;

public sealed class DialogWatcher
{
    private static readonly HashSet<string> WatchedCaptions = new(StringComparer.Ordinal)
    {
        "Microsoft Visual Basic for Applications",
        "Microsoft Excel"
    };

    private readonly IWin32Windows _windows;
    private readonly int _targetProcessId;
    private readonly object _lock = new();
    private CapturedDialog? _captured;
    private Thread? _thread;
    private volatile bool _running;

    public DialogWatcher(IWin32Windows windows, int targetProcessId)
    {
        _windows = windows;
        _targetProcessId = targetProcessId;
    }

    public CapturedDialog? Captured
    {
        get { lock (_lock) { return _captured; } }
    }

    public void Start()
    {
        _running = true;
        _thread = new Thread(PollLoop) { IsBackground = true };
        _thread.Start();
    }

    public void Stop()
    {
        _running = false;
        _thread?.Join();
        _thread = null;
    }

    private void PollLoop()
    {
        while (_running)
        {
            try
            {
                PollOnce();
            }
            catch
            {
                // A malformed window shape shouldn't kill the loop - per spec,
                // per-tick failures are swallowed; only a fatal exception here
                // would escape, and there's none expected in this method.
            }

            Thread.Sleep(100);
        }
    }

    public void PollOnce()
    {
        foreach (var hwnd in _windows.EnumerateTopLevelWindows())
        {
            if (_windows.GetWindowProcessId(hwnd) != _targetProcessId)
            {
                continue;
            }

            var caption = _windows.GetWindowText(hwnd);
            if (!WatchedCaptions.Contains(caption))
            {
                continue;
            }

            var bodyParts = new List<string>();
            var buttonHandle = IntPtr.Zero;

            foreach (var child in _windows.EnumerateChildWindows(hwnd))
            {
                var className = _windows.GetClassName(child);
                if (className == "Static")
                {
                    bodyParts.Add(_windows.GetWindowText(child));
                }
                else if (className == "Button" && buttonHandle == IntPtr.Zero)
                {
                    buttonHandle = child;
                }
            }

            var body = string.Join(" ", bodyParts.Where(s => !string.IsNullOrWhiteSpace(s)));

            lock (_lock)
            {
                _captured = new CapturedDialog(caption, body);
            }

            if (buttonHandle != IntPtr.Zero)
            {
                _windows.Click(buttonHandle);
            }
        }
    }
}
