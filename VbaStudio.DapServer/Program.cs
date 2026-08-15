// VbaStudio.DapServer/Program.cs
using System;
using System.IO;
using System.Text.Json;
using VbaStudio.Core.Excel;
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
}
