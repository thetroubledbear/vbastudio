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

public sealed record CompileResult(bool Success, Diagnostic? Diagnostic);
public sealed record RunResult(bool Success, object? ReturnValue, Diagnostic? Diagnostic);

public sealed class Runner
{
    private const int CompileVbaProjectControlId = 578;

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    private readonly Excel.Application _excel;
    private readonly VBIDE.VBProject _project;
    private readonly VBIDE.VBE _vbe;

    public Runner(Excel.Application excel, VBIDE.VBProject project)
    {
        _excel = excel;
        _project = project;
        _vbe = excel.VBE;
    }

    public CompileResult CompileOnly()
    {
        var diagnostic = ExecuteWithWatcher(() =>
        {
            var commandBars = _excel.CommandBars;
            try
            {
                var control = commandBars.FindControl(Id: CompileVbaProjectControlId);
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
        var diagnostic = ExecuteWithWatcher(() =>
        {
            returnValue = _excel.Run(entryPoint);
        });

        return new RunResult(diagnostic == null, returnValue, diagnostic);
    }

    private Diagnostic? ExecuteWithWatcher(Action blockingCall)
    {
        var processId = GetExcelProcessId();
        var watcher = new DialogWatcher(new Win32Windows(), processId);

        watcher.Start();
        try
        {
            blockingCall();
        }
        finally
        {
            watcher.Stop();
        }

        var captured = watcher.Captured;
        return captured == null ? null : CorrelateDiagnostic(captured);
    }

    private Diagnostic CorrelateDiagnostic(CapturedDialog captured)
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

    private int GetExcelProcessId()
    {
        GetWindowThreadProcessId((IntPtr)_excel.Hwnd, out uint processId);
        return (int)processId;
    }
}
