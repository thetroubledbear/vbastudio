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

        // Anything that looks like a mode but matched none of the branches above (a typo, or a
        // known mode missing its required argument - "stale" with no module name) must NOT fall
        // through into the DAP loop below: that loop blocks on stdin forever, so a malformed
        // invocation would hang instead of failing. Exit 2 distinguishes "bad invocation" from
        // the one-shot modes' exit 1 ("ran, but Excel/COM failed").
        if (args.Length > 0)
        {
            Console.Error.WriteLine(
                $"Unknown mode '{args[0]}'. Expected one of: list, pull, push, stale <moduleName>, " +
                "or no arguments at all to run as a DAP server on stdin/stdout.");
            Environment.Exit(2);
            return;
        }

        RunDapServer();
    }

    // The shared prologue/epilogue every mode below needs: register the COM message filter, attach
    // to the already-running Excel instance, refuse cleanly if there's no active workbook, hand the
    // body the workbook's own directory (the root of this project's <workbookDir>/src convention),
    // and make sure any failure leaves via one stderr line and a non-zero exit rather than an
    // unhandled exception's stack trace - the caller here is VS Code's extension host via
    // child_process, not a human reading a console. The filter is revoked either way.
    private static void WithActiveWorkbook(Action<Excel.Application, Excel.Workbook, string> body)
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
            body(excel, workbook, workbookDir);
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

    // The default, no-arguments mode: a long-lived DAP server speaking on stdin/stdout, unlike the
    // four one-shot modes below. Only the attach-and-validate prologue is shared with them - once
    // DapSession.RunMessageLoop starts, its own error handling takes over and nothing here
    // interferes with it.
    private static void RunDapServer()
    {
        WithActiveWorkbook((excel, workbook, workbookDir) =>
        {
            var shadowPath = Path.Combine(workbookDir, "build", "shadow.xlsm");
            Directory.CreateDirectory(Path.GetDirectoryName(shadowPath)!);

            using var input = Console.OpenStandardInput();
            using var output = Console.OpenStandardOutput();

            var session = new DapSession(excel, workbook, shadowPath, output);
            session.RunMessageLoop(input);
        });
    }

    // Separate, short-lived, one-shot code path from the DAP message loop above - not a DAP
    // request, never touches DapSession. Prints one JSON line to stdout and exits. Any failure
    // (Excel not running, no active workbook, COM error) writes one line to stderr and exits
    // non-zero - see WithActiveWorkbook.
    private static void RunList()
    {
        WithActiveWorkbook((excel, workbook, workbookDir) =>
        {
            var access = new ExcelVbaProjectAccess(workbook.VBProject);
            var modules = access.ReadAll();
            var result = ModuleListBuilder.Build(workbook.FullName, modules);
            Console.WriteLine(JsonSerializer.Serialize(result));
        });
    }

    // Same one-shot shape as RunList - attaches, runs SyncEngine.Pull() (exports every module to
    // <workbookDir>/src/<Kind>/<name>.ext, the same convention DapSession.HandleLaunch and
    // RunStale below both assume), prints a success marker, exits. The resolved srcDir is part of
    // that marker because it is derived from the WORKBOOK's directory, which is not guaranteed to
    // be the folder open in VS Code - the extension shows this path back to the user so a pull
    // into a directory they aren't looking at is at least visible rather than silent.
    private static void RunPull()
    {
        WithActiveWorkbook((excel, workbook, workbookDir) =>
        {
            var srcDir = Path.Combine(workbookDir, "src");
            var access = new ExcelVbaProjectAccess(workbook.VBProject);
            var syncEngine = new SyncEngine(access, new FileSystem(), srcDir);
            var result = syncEngine.Pull();
            Console.WriteLine(JsonSerializer.Serialize(new
            {
                success = true,
                srcDir,
                written = result.Written,
                deleted = result.Deleted,
                conflicts = result.Conflicts
            }));
        });
    }

    // Same one-shot shape as RunPull - runs SyncEngine.Push() (writes every changed disk file
    // back into the live project). Push's own IsMacroRunning guard - currently a hardcoded false
    // stub, see ExcelVbaProjectAccess.cs - would throw InvalidOperationException here if it ever
    // becomes real; WithActiveWorkbook's catch already surfaces that cleanly as a stderr message.
    private static void RunPush()
    {
        WithActiveWorkbook((excel, workbook, workbookDir) =>
        {
            var srcDir = Path.Combine(workbookDir, "src");
            var access = new ExcelVbaProjectAccess(workbook.VBProject);
            var syncEngine = new SyncEngine(access, new FileSystem(), srcDir);
            var result = syncEngine.Push();
            Console.WriteLine(JsonSerializer.Serialize(new
            {
                success = true,
                written = result.Written,
                deleted = result.Deleted,
                conflicts = result.Conflicts
            }));
        });
    }

    // Reads the named module's current Excel code and its corresponding disk file (if any), and
    // reports whether they differ via the pure StaleChecker.IsStale (Core, unit-tested). A module
    // name Excel doesn't have (e.g. renamed since the last pull) is reported as stale rather than
    // a separate error path, matching the spec's error-handling section.
    private static void RunStale(string moduleName)
    {
        WithActiveWorkbook((excel, workbook, workbookDir) =>
        {
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
        });
    }
}
