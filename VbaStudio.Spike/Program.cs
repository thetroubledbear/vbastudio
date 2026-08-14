using System;
using System.IO.Abstractions;
using VbaStudio.Core.Excel;
using VbaStudio.Core.Sync;
using VbaStudio.Core.Testing;
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
            else if (args.Length > 0 && args[0] == "test")
            {
                RunTests();
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
        var runner = new Runner(excel, project, Console.WriteLine);

        // No separate compile step - Run() implicitly compiles the whole project first. See
        // Runner.Run()'s comment for why: it is the only compile-check mechanism proven reliable.
        Console.WriteLine($"Running {entryPoint}...");
        var start = DateTime.UtcNow;
        RunResult run;
        try
        {
            run = runner.Run(entryPoint);
        }
        catch (InvalidOperationException ex)
        {
            // EnsureTargetProjectIsActive / EnsureProjectIsInDesignMode refusing a precondition -
            // no COM call was attempted, nothing to clean up. A stack trace here would be noise;
            // the message already says exactly what to do.
            Console.WriteLine($"REFUSED: {ex.Message}");
            return;
        }
        var elapsed = DateTime.UtcNow - start;
        Console.WriteLine($"Finished in {elapsed.TotalMilliseconds:F0}ms");

        if (!run.Success)
        {
            var d = run.Diagnostic!;
            Console.WriteLine($"ERROR: module={d.Module ?? "?"} line={d.Line?.ToString() ?? "?"} message={d.Message}");
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

    private static void RunTests()
    {
        var excel = (Excel.Application)ComHelpers.GetRunningInstance("Excel.Application");
        var project = excel.ActiveWorkbook.VBProject;
        var runner = new Runner(excel, project, Console.WriteLine);
        var testRunner = new TestRunner(runner);

        var tests = TestDiscovery.DiscoverTests(new FileSystem(), "src");
        if (tests.Count == 0)
        {
            Console.WriteLine("0 tests found.");
            return;
        }

        Console.WriteLine($"Discovered {tests.Count} test(s) (from src/ on disk - Push first if you've edited them).");

        var results = testRunner.RunAll(tests);

        var passed = 0;
        var failed = 0;
        var skipped = 0;
        foreach (var result in results)
        {
            if (result.Skipped)
            {
                // Never ran - Runner's preconditions refused it (project stuck out of design mode).
                // Not a red assertion, so it must not inflate the failure count.
                skipped++;
                Console.WriteLine($"SKIP {result.Test.QualifiedName}: {result.FailureMessage}");
            }
            else if (result.Passed)
            {
                passed++;
                Console.WriteLine($"PASS {result.Test.QualifiedName} ({result.Duration.TotalMilliseconds:F0}ms)");
            }
            else
            {
                failed++;
                Console.WriteLine($"FAIL {result.Test.QualifiedName}: {result.FailureMessage} ({result.Duration.TotalMilliseconds:F0}ms)");
            }
        }

        Console.WriteLine($"{passed} passed, {failed} failed, {skipped} skipped, {results.Count} total");
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
