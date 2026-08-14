// VbaStudio.Interop/Runner.cs
using System;
using System.Runtime.InteropServices;
using VbaStudio.Core.Excel;
using VbaStudio.Core.Model;
using VbaStudio.Core.Win32;
using Excel = Microsoft.Office.Interop.Excel;
// tlbimp generates VBE interop types directly in namespace VBIDE, not Microsoft.Vbe.Interop; aliasing fails with CS0234
using VBIDE;

namespace VbaStudio.Interop;

public sealed record RunResult(bool Success, object? ReturnValue, Diagnostic? Diagnostic);

public sealed class Runner
{
    // Only this caption means "VBA compile/runtime error". A "Microsoft Excel" dialog during
    // Application.Run is typically a MsgBox from the macro itself - dismissed, not a failure.
    private const string VbaErrorCaption = "Microsoft Visual Basic for Applications";

    private readonly Excel.Application _excel;
    private readonly VBIDE.VBProject _project;
    private readonly VBIDE.VBE _vbe;
    private readonly Action<string>? _log;
    private readonly Win32Windows _windows = new();

    public Runner(Excel.Application excel, VBIDE.VBProject project, Action<string>? log = null)
    {
        _excel = excel;
        _project = project;
        _vbe = excel.VBE;
        _log = log;
    }

    public RunResult Run(string entryPoint)
    {
        EnsureTargetProjectIsActive();
        EnsureProjectIsInDesignMode();

        // There is no reliable side-effect-free compile check via automation in this VBE version.
        // Two approaches were tried and both failed against real Excel:
        //  1. The VBE's own "Compile VBAProject" command (control id 578) does not reliably
        //     surface errors - Execute() can return cleanly in under 100ms on code that does not
        //     compile, with no dialog ever appearing.
        //  2. Injecting a throwaway no-op procedure and Application.Run-ing it to force a
        //     whole-project compile without executing real code: when the error is in a
        //     *different* module than the one being run, VBA does not show a dismissible dialog
        //     at all - it drops the VBE straight into break mode on the broken module, with
        //     nothing for a window-enumeration-based watcher to click. Confirmed against a live
        //     Excel session; recovering from it required manually closing the workbook.
        // The one mechanism proven reliable, every time, is calling Application.Run directly on
        // the actual target entry point: VBA's implicit "compile the whole project before
        // executing" step then surfaces a real, dismissible "Compile error" dialog for that
        // module. A compile failure and a genuine runtime error surface through the identical
        // dialog, and M2's contract only promises {module, line, message} for either - not which
        // kind it was - so Run() alone satisfies the exit gate without a separate compile step.
        object? returnValue = null;
        COMException? runException = null;
        var captured = ExecuteWithWatcher(() =>
        {
            try
            {
                returnValue = _excel.Run(entryPoint);
            }
            catch (COMException ex)
            {
                // Application.Run does not always report a compile/runtime failure via a dialog -
                // confirmed empirically: it can also throw a raw COMException straight back to the
                // automation caller (0x800ADF09 observed against real Excel) with no window ever
                // appearing for DialogWatcher to see. This must not crash the process uncaught.
                runException = ex;
            }
        });

        // Checked regardless of outcome: a stuck break-mode state after a failure is exactly what
        // let today's crashes cascade (a second Run() call, chained onto a project VBE never
        // finished recovering from). This does not fix that state - VBProject.Mode has no
        // documented setter to force one - but a caller ignoring the warning is now ignoring a
        // clear signal, not walking into a silent trap.
        LogIfNotBackInDesignMode();

        if (captured != null && captured.Caption == VbaErrorCaption)
        {
            var diagnostic = CorrelateDiagnostic(captured.Body);
            return new RunResult(false, null, diagnostic);
        }

        if (runException != null)
        {
            var diagnostic = CorrelateDiagnostic(runException.Message);
            return new RunResult(false, null, diagnostic);
        }

        // No dialog, no exception, or a non-VBA-error dialog (e.g. the macro's own MsgBox,
        // already dismissed by the watcher): the macro ran to completion, report its return value.
        return new RunResult(true, returnValue, null);
    }

    /// <summary>
    /// Arms the dialog watcher, makes the blocking COM call, tears the watcher back down, and
    /// returns whatever dialog was captured. Interpreting the capture is the caller's job.
    /// </summary>
    private CapturedDialog? ExecuteWithWatcher(Action blockingCall)
    {
        var processId = _windows.GetWindowProcessId((IntPtr)_excel.Hwnd);
        var watcher = new DialogWatcher(_windows, processId, _log);

        watcher.Start();
        try
        {
            blockingCall();
        }
        finally
        {
            watcher.Stop();

            if (watcher.FatalError != null)
            {
                _log?.Invoke($"Runner: dialog watcher died with a fatal error: {watcher.FatalError}");
            }
        }

        return watcher.Captured;
    }

    private Diagnostic CorrelateDiagnostic(string message)
    {
        // The message is the one thing we already know (from a dismissed dialog's body, or a
        // caught COMException's own message); a COM failure while correlating it to a module/line
        // must degrade to a message-only diagnostic, never lose the message.
        try
        {
            var pane = _vbe.ActiveCodePane;
            if (pane == null)
            {
                return new Diagnostic(Module: null, Line: null, message);
            }

            try
            {
                var module = pane.CodeModule;
                try
                {
                    var component = module.Parent;
                    try
                    {
                        var moduleName = component.Name;
                        pane.GetSelection(out int startLine, out _, out _, out _);
                        return new Diagnostic(moduleName, startLine, message);
                    }
                    finally
                    {
                        ComRelease.Release(component);
                    }
                }
                finally
                {
                    ComRelease.Release(module);
                }
            }
            finally
            {
                ComRelease.Release(pane);
            }
        }
        catch (Exception ex)
        {
            _log?.Invoke($"Runner: could not correlate the failure to a module/line: {ex.Message}");
            return new Diagnostic(Module: null, Line: null, message);
        }
    }

    /// <summary>
    /// Compile and correlation both act on the VBE's active project, so refuse to run when that
    /// isn't the project the caller named - otherwise we'd silently compile someone else's code.
    /// VBIDE offers no clean way to activate a project programmatically, so this validates only.
    /// </summary>
    private void EnsureTargetProjectIsActive()
    {
        var expectedName = _project.Name;
        var activeProject = _vbe.ActiveVBProject;
        try
        {
            var actualName = activeProject?.Name;
            if (!string.Equals(actualName, expectedName, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    $"The target VBA project '{expectedName}' is not the active project in the VBE " +
                    $"(active project: '{actualName ?? "<none>"}'). Select it in Excel or the VBE before compiling or running.");
            }
        }
        finally
        {
            ComRelease.Release(activeProject);
        }
    }

    /// <summary>
    /// Refuses to start a run while the project is mid-execution or paused in break mode from a
    /// prior call - the exact precondition Rubberduck's own test runner checks before running
    /// anything (VBProject.Mode, a documented Microsoft.Vbe.Interop property; not the reverse-
    /// engineered internals their execution engine itself needs). Chaining a second Run() onto a
    /// project VBE never finished recovering from is what turned today's failures into crashes.
    /// </summary>
    private void EnsureProjectIsInDesignMode()
    {
        var mode = _project.Mode;
        if (mode != vbext_VBAMode.vbext_vm_Design)
        {
            throw new InvalidOperationException(
                $"The target VBA project '{_project.Name}' is not in design mode (current mode: {mode}). " +
                "It is likely still mid-execution or paused in break mode from a prior failure - reset it " +
                "in Excel (Run > Reset, or close and reopen the workbook) before running again.");
        }
    }

    private void LogIfNotBackInDesignMode()
    {
        var mode = _project.Mode;
        if (mode != vbext_VBAMode.vbext_vm_Design)
        {
            _log?.Invoke(
                $"Runner: project '{_project.Name}' did not return to design mode after Run() " +
                $"(current mode: {mode}). The next call may fail or crash Excel until this is reset by hand.");
        }
    }
}
