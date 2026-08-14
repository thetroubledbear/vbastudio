// VbaStudio.DapServer/Program.cs
using System;
using System.IO;
using VbaStudio.Core.Dap;
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

            while (true)
            {
                var request = DapProtocol.ReadRequest(input);
                if (request == null)
                {
                    break;
                }

                session.HandleRequest(request);
            }
        }
        finally
        {
            ExcelMessageFilter.Revoke();
        }
    }
}
