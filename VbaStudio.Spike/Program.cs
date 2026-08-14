using System;
using System.IO.Abstractions;
using VbaStudio.Core.Excel;
using VbaStudio.Core.Sync;
using VbaStudio.Interop;
using Excel = Microsoft.Office.Interop.Excel;
// tlbimp generates VBE interop types directly in namespace VBIDE, not Microsoft.Vbe.Interop; aliasing fails with CS0234
using VBIDE;

namespace VbaStudio.Spike;

internal class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        ExcelMessageFilter.Register();
        try
        {
            if (args.Length > 0 && args[0] == "synctest")
            {
                RunSyncTest();
            }
            else if (args.Length > 0 && args[0] == "run")
            {
                var entryPoint = args.Length > 1 ? args[1] : "modMain.Main";
                RunCompileAndRun(entryPoint);
            }
            else
            {
                RunSpike();
            }
        }
        finally
        {
            ExcelMessageFilter.Revoke();
        }
    }

    private static void RunCompileAndRun(string entryPoint)
    {
        var excel = (Excel.Application)ComHelpers.GetRunningInstance("Excel.Application");
        var project = excel.ActiveWorkbook.VBProject;
        var runner = new Runner(excel, project);

        Console.WriteLine("Compiling...");
        var compileStart = DateTime.UtcNow;
        var compile = runner.CompileOnly();
        var compileElapsed = DateTime.UtcNow - compileStart;
        Console.WriteLine($"Compile finished in {compileElapsed.TotalMilliseconds:F0}ms");

        if (!compile.Success)
        {
            var d = compile.Diagnostic!;
            Console.WriteLine($"COMPILE ERROR: module={d.Module ?? "?"} line={d.Line?.ToString() ?? "?"} message={d.Message}");
            return;
        }

        Console.WriteLine($"Compile clean. Running {entryPoint}...");
        var run = runner.Run(entryPoint);
        if (!run.Success)
        {
            var d = run.Diagnostic!;
            Console.WriteLine($"RUNTIME ERROR: module={d.Module ?? "?"} line={d.Line?.ToString() ?? "?"} message={d.Message}");
            return;
        }

        Console.WriteLine($"Run complete. Return value: {run.ReturnValue}");
    }

    private static void RunSyncTest()
    {
        var excel = (Excel.Application)ComHelpers.GetRunningInstance("Excel.Application");
        var project = excel.ActiveWorkbook.VBProject;
        var access = new ExcelVbaProjectAccess(project);
        var sync = new SyncEngine(access, new FileSystem(), "src");

        Console.WriteLine("Pull #1...");
        sync.Pull();
        Console.WriteLine("Push (unchanged files)...");
        sync.Push();
        Console.WriteLine("Pull #2...");
        sync.Pull();
        Console.WriteLine("Done. Diff the 'src' directory between runs with git or a hash tool to confirm byte-identical output.");
    }

    private static void RunSpike()
    {
        var excel = (Excel.Application)ComHelpers.GetRunningInstance("Excel.Application");
        Console.WriteLine($"Attached to Excel {excel.Version}");

        Excel.Workbook book = excel.ActiveWorkbook;
        Console.WriteLine($"Workbook: {book.Name}");

        VBProject project;
        try
        {
            project = book.VBProject;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Cannot reach the VBProject.");
            Console.WriteLine("Trust Center > Macro Settings > Trust access to the VBA project object model.");
            Console.WriteLine($"({ex.Message})");
            return;
        }

        foreach (VBComponent component in project.VBComponents)
        {
            int lines = component.CodeModule.CountOfLines;
            Console.WriteLine($"{component.Name,-30} type={(int)component.Type,-4} lines={lines}");
        }
    }
}
