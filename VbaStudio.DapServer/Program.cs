// VbaStudio.DapServer/Program.cs
using System;
using System.IO;
using System.IO.Abstractions;
using System.Linq;
using System.Text.Json;
using VbaStudio.Core.Excel;
using VbaStudio.Core.Model;
using VbaStudio.Core.Sync;
using VbaStudio.Core.Tooling;
using VbaStudio.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace VbaStudio.DapServer;

internal class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        if (args.Length > 0 && string.Equals(args[0], "list", StringComparison.OrdinalIgnoreCase))
        {
            RunList();
            return;
        }

        if (args.Length > 0 && string.Equals(args[0], "pull", StringComparison.OrdinalIgnoreCase))
        {
            RunPull();
            return;
        }

        if (args.Length > 0 && string.Equals(args[0], "push", StringComparison.OrdinalIgnoreCase))
        {
            RunPush();
            return;
        }

        if (args.Length > 1 && string.Equals(args[0], "stale", StringComparison.OrdinalIgnoreCase))
        {
            RunStale(args[1]);
            return;
        }

        ExcelMessageFilter.Register();
        try
        {
            var excel = (Excel.Application)ComHelpers.GetRunningInstance("Excel.Application");
            var workbook = excel.ActiveWorkbook;

            var workbookDir = Path.GetDirectoryName(workbook.FullName) ?? ".";
            var shadowPath = Path.Combine(workbookDir, "build", "shadow.xlsm");
            Directory.CreateDirectory(Path.GetDirectoryName(shadowPath)!);

            using var input = Console.OpenStandardInput();
            using var output = Console.OpenStandardOutput();

            var session = new DapSession(excel, workbook, shadowPath, output);
            session.RunMessageLoop(input);
        }
        finally
        {
            ExcelMessageFilter.Revoke();
        }
    }

    // Separate, short-lived, one-shot code path from the DAP message loop above - not a DAP
    // request, never touches DapSession. Prints one JSON line to stdout and exits. Any failure
    // (Excel not running, no active workbook, COM error) writes one line to stderr and exits
    // non-zero rather than letting an exception's stack trace reach the caller - the caller here
    // is VS Code's extension host via child_process, not a human reading a console.
    private static void RunList()
    {
        ExcelMessageFilter.Register();
        try
        {
            var excel = (Excel.Application)ComHelpers.GetRunningInstance("Excel.Application");
            var workbook = excel.ActiveWorkbook;
            if (workbook == null)
            {
                Console.Error.WriteLine("No active workbook in the running Excel instance.");
                Environment.Exit(1);
                return;
            }

            var access = new ExcelVbaProjectAccess(workbook.VBProject);
            var modules = access.ReadAll();
            var result = ModuleListBuilder.Build(workbook.FullName, modules);
            Console.WriteLine(JsonSerializer.Serialize(result));
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            Environment.Exit(1);
        }
        finally
        {
            ExcelMessageFilter.Revoke();
        }
    }

    // Same one-shot shape as RunList - attaches, runs SyncEngine.Pull() (exports every module to
    // <workbookDir>/src/<Kind>/<name>.ext, the same convention DapSession.HandleLaunch and
    // RunStale below both assume), prints a success marker, exits.
    private static void RunPull()
    {
        ExcelMessageFilter.Register();
        try
        {
            var excel = (Excel.Application)ComHelpers.GetRunningInstance("Excel.Application");
            var workbook = excel.ActiveWorkbook;
            if (workbook == null)
            {
                Console.Error.WriteLine("No active workbook in the running Excel instance.");
                Environment.Exit(1);
                return;
            }

            var workbookDir = Path.GetDirectoryName(workbook.FullName) ?? ".";
            var srcDir = Path.Combine(workbookDir, "src");
            var access = new ExcelVbaProjectAccess(workbook.VBProject);
            var syncEngine = new SyncEngine(access, new FileSystem(), srcDir);
            syncEngine.Pull();
            Console.WriteLine("{\"success\":true}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            Environment.Exit(1);
        }
        finally
        {
            ExcelMessageFilter.Revoke();
        }
    }

    // Same one-shot shape as RunPull - runs SyncEngine.Push() (writes every changed disk file
    // back into the live project). Push's own IsMacroRunning guard - currently a hardcoded false
    // stub, see ExcelVbaProjectAccess.cs - would throw InvalidOperationException here if it ever
    // becomes real; the catch below already surfaces that cleanly as a stderr message.
    private static void RunPush()
    {
        ExcelMessageFilter.Register();
        try
        {
            var excel = (Excel.Application)ComHelpers.GetRunningInstance("Excel.Application");
            var workbook = excel.ActiveWorkbook;
            if (workbook == null)
            {
                Console.Error.WriteLine("No active workbook in the running Excel instance.");
                Environment.Exit(1);
                return;
            }

            var workbookDir = Path.GetDirectoryName(workbook.FullName) ?? ".";
            var srcDir = Path.Combine(workbookDir, "src");
            var access = new ExcelVbaProjectAccess(workbook.VBProject);
            var syncEngine = new SyncEngine(access, new FileSystem(), srcDir);
            syncEngine.Push();
            Console.WriteLine("{\"success\":true}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            Environment.Exit(1);
        }
        finally
        {
            ExcelMessageFilter.Revoke();
        }
    }

    // Reads the named module's current Excel code and its corresponding disk file (if any), and
    // reports whether they differ via the pure StaleChecker.IsStale (Core, unit-tested). A module
    // name Excel doesn't have (e.g. renamed since the last pull) is reported as stale rather than
    // a separate error path, matching the spec's error-handling section.
    private static void RunStale(string moduleName)
    {
        ExcelMessageFilter.Register();
        try
        {
            var excel = (Excel.Application)ComHelpers.GetRunningInstance("Excel.Application");
            var workbook = excel.ActiveWorkbook;
            if (workbook == null)
            {
                Console.Error.WriteLine("No active workbook in the running Excel instance.");
                Environment.Exit(1);
                return;
            }

            var workbookDir = Path.GetDirectoryName(workbook.FullName) ?? ".";
            var access = new ExcelVbaProjectAccess(workbook.VBProject);
            var modules = access.ReadAll();
            var targetModule = modules.FirstOrDefault(
                m => string.Equals(m.Name, moduleName, StringComparison.OrdinalIgnoreCase));

            if (targetModule == null)
            {
                Console.WriteLine("{\"stale\":true}");
                return;
            }

            var path = Path.Combine(workbookDir, "src", targetModule.Kind.SourceFolder(), targetModule.FileName);
            string? diskContent = File.Exists(path)
                ? File.ReadAllText(path, targetModule.Kind.SourceEncoding())
                : null;

            var stale = StaleChecker.IsStale(diskContent, targetModule.Code);
            Console.WriteLine(JsonSerializer.Serialize(new { stale }));
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            Environment.Exit(1);
        }
        finally
        {
            ExcelMessageFilter.Revoke();
        }
    }
}
