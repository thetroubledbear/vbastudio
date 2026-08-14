// VbaStudio.Interop/Runner.cs
using System;
using VbaStudio.Core.Excel;
using VbaStudio.Core.Model;
using VbaStudio.Core.Win32;
using Excel = Microsoft.Office.Interop.Excel;
// tlbimp generates VBE interop types directly in namespace VBIDE, not Microsoft.Vbe.Interop; aliasing fails with CS0234
using VBIDE;

namespace VbaStudio.Interop;

public sealed record CompileResult(bool Success, Diagnostic? Diagnostic);
public sealed record RunResult(bool Success, object? ReturnValue, Diagnostic? Diagnostic);

public sealed class Runner
{
    private const int CompileVbaProjectControlId = 578;

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

    public CompileResult CompileOnly()
    {
        EnsureTargetProjectIsActive();

        var captured = ExecuteWithWatcher(() =>
        {
            // Control id 578 is a VBE command ("Debug > Compile VBAProject"), so it must be
            // resolved against the VBE's command bars, not the host application's.
            var commandBars = _vbe.CommandBars;
            try
            {
                var control = commandBars.FindControl(Id: CompileVbaProjectControlId);
                if (control == null)
                {
                    throw new InvalidOperationException(
                        $"VBE command {CompileVbaProjectControlId} (Compile VBAProject) not found. " +
                        "Is 'Trust access to the VBA project object model' enabled?");
                }

                try
                {
                    control.Execute();
                }
                finally
                {
                    ComRelease.Release(control);
                }
            }
            finally
            {
                ComRelease.Release(commandBars);
            }
        });

        // Any dialog captured during the compile step means the compile failed.
        var diagnostic = captured == null ? null : CorrelateDiagnostic(captured);
        return new CompileResult(diagnostic == null, diagnostic);
    }

    public RunResult Run(string entryPoint)
    {
        var compile = CompileOnly();
        if (!compile.Success)
        {
            return new RunResult(false, null, compile.Diagnostic);
        }

        object? returnValue = null;
        var captured = ExecuteWithWatcher(() =>
        {
            returnValue = _excel.Run(entryPoint);
        });

        if (captured != null && captured.Caption == VbaErrorCaption)
        {
            var diagnostic = CorrelateDiagnostic(captured);
            return new RunResult(false, null, diagnostic);
        }

        // No dialog, or a non-VBA-error dialog (e.g. the macro's own MsgBox, already
        // dismissed by the watcher): the macro ran to completion, so report its return value.
        return new RunResult(true, returnValue, null);
    }

    /// <summary>
    /// Arms the dialog watcher, makes the blocking COM call, tears the watcher back down, and
    /// returns whatever dialog was captured. Interpreting the capture is the caller's job -
    /// CompileOnly() and Run() disagree about what a captured dialog means.
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

    private Diagnostic CorrelateDiagnostic(CapturedDialog captured)
    {
        // The captured body is the one thing we already know; a COM failure while correlating
        // it to a module/line must degrade to a message-only diagnostic, never lose the message.
        try
        {
            var pane = _vbe.ActiveCodePane;
            if (pane == null)
            {
                return new Diagnostic(Module: null, Line: null, captured.Body);
            }

            var module = pane.CodeModule;
            try
            {
                var component = module.Parent;
                try
                {
                    var moduleName = component.Name;
                    pane.GetSelection(out int startLine, out _, out _, out _);
                    return new Diagnostic(moduleName, startLine, captured.Body);
                }
                finally
                {
                    ComRelease.Release(component);
                }
            }
            finally
            {
                ComRelease.Release(module, pane);
            }
        }
        catch (Exception ex)
        {
            _log?.Invoke($"Runner: could not correlate the captured dialog to a module/line: {ex.Message}");
            return new Diagnostic(Module: null, Line: null, captured.Body);
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
}
