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
    // Only this caption means "VBA compile/runtime error". A "Microsoft Excel" dialog during
    // Application.Run is typically a MsgBox from the macro itself - dismissed, not a failure.
    private const string VbaErrorCaption = "Microsoft Visual Basic for Applications";

    // The name is deliberately unusual, not fully collision-proof - VBIDE auto-renames on
    // collision (Add() returns whatever name it actually assigned; we always reference that
    // returned object, never a name lookup), so a clash only costs an odd module name, nothing
    // else.
    private const string CompileCheckProcedureName = "VbaStudioCompileCheck";

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

        // The VBE's own "Compile VBAProject" command (control id 578) does not reliably surface
        // errors when driven by automation - confirmed empirically: Execute() can return cleanly
        // in under 100ms on code that does not compile, with no dialog ever appearing. The one
        // mechanism proven reliable is Application.Run's own implicit "compile the whole project
        // before executing anything" step. To get a side-effect-free compile check out of that,
        // inject a throwaway no-op procedure, Run it, then always remove it again.
        //
        // Every phase below runs under its own watcher, not just the Run() phase - confirmed the
        // hard way. VBComponents.Add() against a project with a just-edited, not-yet-validated
        // module (the exact CompileOnly use case) can itself pop a modal: "This action will reset
        // your project, proceed anyway?" - same caption/class as a compile error, but not one.
        // With no watcher running yet at that point, it blocked indefinitely. A dialog from setup
        // or cleanup gets dismissed but is not diagnostic-worthy; only a dialog during the actual
        // compile-check Run() means the target code failed to compile - so each phase gets its
        // own fresh watcher, and only the middle one's capture becomes the returned diagnostic.
        VBComponent? checkComponent = null;
        ExecuteWithWatcher(() =>
        {
            var components = _project.VBComponents;
            try
            {
                checkComponent = components.Add(vbext_ComponentType.vbext_ct_StdModule);
            }
            finally
            {
                ComRelease.Release(components);
            }

            var codeModule = checkComponent.CodeModule;
            try
            {
                codeModule.AddFromString($"Sub {CompileCheckProcedureName}()\r\nEnd Sub");
            }
            finally
            {
                ComRelease.Release(codeModule);
            }
        });

        if (checkComponent == null)
        {
            throw new InvalidOperationException("Failed to create the compile-check module.");
        }

        var qualifiedName = $"{checkComponent.Name}.{CompileCheckProcedureName}";
        var captured = ExecuteWithWatcher(() =>
        {
            _excel.Run(qualifiedName);
        });

        var componentsForRemove = _project.VBComponents;
        try
        {
            ExecuteWithWatcher(() => componentsForRemove.Remove(checkComponent));
        }
        finally
        {
            ComRelease.Release(componentsForRemove);
            ComRelease.Release(checkComponent);
        }

        var diagnostic = captured == null ? null : CorrelateDiagnostic(captured);
        return new CompileResult(diagnostic == null, diagnostic);
    }

    public RunResult Run(string entryPoint)
    {
        EnsureTargetProjectIsActive();

        // No separate compile step: Application.Run compiles the whole project as a side effect
        // before executing anything, which is the one mechanism proven reliable (see CompileOnly).
        // A compile failure and a genuine runtime error surface through the identical dialog, and
        // M2's contract only promises {module, line, message} for either - not which kind it was.
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
                        return new Diagnostic(moduleName, startLine, captured.Body);
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
