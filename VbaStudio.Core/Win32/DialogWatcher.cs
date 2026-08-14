using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace VbaStudio.Core.Win32;

public sealed class DialogWatcher
{
    // Standard Windows dialog box class. Every VBA/Excel modal we target is a #32770;
    // Excel's own frame window (XLMAIN) can share the "Microsoft Excel" caption, so the
    // class check keeps us from enumerating and clicking inside the main window.
    private const string DialogClassName = "#32770";

    private static readonly TimeSpan StopTimeout = TimeSpan.FromSeconds(2);

    // Click() posts asynchronously (never blocks - a hung Excel must not hang this thread).
    // This bounds how long a single poll tick waits for the dialog to actually disappear before
    // moving on. Without it, the next operation's watcher can catch the same window still
    // closing and misreport it as a fresh failure - the dismissal race this class exists to close.
    private static readonly TimeSpan DismissConfirmTimeout = TimeSpan.FromSeconds(1);

    // Confirmed live (2026-08-14): a VBA *runtime*-error dialog (e.g. from Err.Raise) is captioned
    // "Microsoft Visual Basic" - without "for Applications" - a different caption than the
    // compile-error dialog. Both must be watched or a runtime error's dialog sits undismissed
    // forever, wedging the VBE and failing every subsequent Run() call.
    private static readonly HashSet<string> WatchedCaptions = new(StringComparer.Ordinal)
    {
        "Microsoft Visual Basic for Applications",
        "Microsoft Visual Basic",
        "Microsoft Excel"
    };

    private readonly IWin32Windows _windows;
    private readonly int _targetProcessId;
    private readonly Action<string>? _log;
    private readonly object _lock = new();
    private CapturedDialog? _captured;
    private Thread? _thread;
    private volatile bool _running;
    private volatile Exception? _fatalError;

    public DialogWatcher(IWin32Windows windows, int targetProcessId, Action<string>? log = null)
    {
        _windows = windows;
        _targetProcessId = targetProcessId;
        _log = log;
    }

    public CapturedDialog? Captured
    {
        get { lock (_lock) { return _captured; } }
    }

    /// <summary>
    /// Set if an exception escaped the poll loop and killed the watcher thread. A dead
    /// watcher means nothing will dismiss the dialog, so this is a diagnosis aid the
    /// runner can inspect after its blocking call returns - not a recovery mechanism.
    /// </summary>
    public Exception? FatalError => _fatalError;

    public void Start()
    {
        _running = true;
        _thread = new Thread(PollLoop) { IsBackground = true };
        _thread.Start();
    }

    public void Stop()
    {
        _running = false;

        var thread = _thread;
        if (thread != null && !thread.Join(StopTimeout))
        {
            // Accepted degraded state: the thread is a background thread and touches no COM,
            // so we let it be rather than blocking the runner thread forever.
            _log?.Invoke($"DialogWatcher: watcher thread did not stop within {StopTimeout.TotalSeconds:F0}s; leaving it running in the background.");
        }

        _thread = null;
    }

    private void PollLoop()
    {
        try
        {
            while (_running)
            {
                try
                {
                    PollOnce();
                }
                catch (Exception ex)
                {
                    // A malformed window shape shouldn't kill the loop - per-tick failures
                    // are swallowed, but logged rather than silently discarded.
                    _log?.Invoke($"DialogWatcher: poll tick failed: {ex}");
                }

                Thread.Sleep(100);
            }
        }
        catch (Exception ex)
        {
            _fatalError = ex;
            _log?.Invoke($"DialogWatcher: fatal error, poll loop stopped: {ex}");
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

            if (_windows.GetClassName(hwnd) != DialogClassName)
            {
                continue;
            }

            var bodyParts = new List<string>();
            var firstButton = IntPtr.Zero;
            var okButton = IntPtr.Zero;
            var endButton = IntPtr.Zero;

            foreach (var child in _windows.EnumerateChildWindows(hwnd))
            {
                var className = _windows.GetClassName(child);
                if (className == "Static")
                {
                    bodyParts.Add(_windows.GetWindowText(child));
                }
                else if (className == "Button")
                {
                    if (firstButton == IntPtr.Zero)
                    {
                        firstButton = child;
                    }

                    // Children come back in Z-order, not OK-first order, and the VBA
                    // compile-error dialog also has a Help button - clicking Help would
                    // leave the dialog up, which is the hang this watcher exists to prevent.
                    var childText = _windows.GetWindowText(child);
                    if (okButton == IntPtr.Zero && (childText == "OK" || childText == "&OK"))
                    {
                        okButton = child;
                    }

                    // The runtime-error dialog ("Microsoft Visual Basic" caption) has no OK
                    // button - its set is Continue/End/Debug/Help, with Continue first by
                    // Z-order. Falling back to firstButton there would click Continue (resumes
                    // into undefined further execution) instead of ending the run cleanly.
                    if (endButton == IntPtr.Zero && (childText == "End" || childText == "&End"))
                    {
                        endButton = child;
                    }
                }
            }

            var body = string.Join(" ", bodyParts.Where(s => !string.IsNullOrWhiteSpace(s)));

            lock (_lock)
            {
                _captured = new CapturedDialog(caption, body);
            }

            var buttonHandle = okButton != IntPtr.Zero ? okButton
                : endButton != IntPtr.Zero ? endButton
                : firstButton;

            // Log every caption match - known or unrecognized shape - before dismissing.
            _log?.Invoke(
                $"DialogWatcher: matched dialog caption=\"{caption}\" body=\"{body}\" " +
                (buttonHandle == IntPtr.Zero ? "button=<none found, not dismissing>" : $"button=0x{buttonHandle.ToInt64():X}"));

            if (buttonHandle != IntPtr.Zero)
            {
                _windows.Click(buttonHandle);

                if (!_windows.WaitForWindowClosed(hwnd, DismissConfirmTimeout))
                {
                    _log?.Invoke(
                        $"DialogWatcher: dialog did not confirm closed within {DismissConfirmTimeout.TotalSeconds:F0}s " +
                        "after dismissal - a later operation may catch it as stale.");
                }
            }
        }
    }
}
