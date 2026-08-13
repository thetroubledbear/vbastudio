using System;
using VbaStudio.Core.Excel;
using Excel = Microsoft.Office.Interop.Excel;
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
            RunSpike();
        }
        finally
        {
            ExcelMessageFilter.Revoke();
        }
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
