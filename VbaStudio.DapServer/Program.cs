// VbaStudio.DapServer/Program.cs
using System;
using System.IO;
using VbaStudio.Core.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace VbaStudio.DapServer;

internal class Program
{
    [STAThread]
    static void Main(string[] args)
    {
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
}
