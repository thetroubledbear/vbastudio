using Excel = Microsoft.Office.Interop.Excel;

namespace VbaStudio.Interop;

public static class ExcelShadowWorkbook
{
    public static Excel.Workbook CreateFromOpen(Excel.Workbook workbook, string shadowPath)
    {
        workbook.SaveCopyAs(shadowPath);
        return workbook.Application.Workbooks.Open(shadowPath);
    }
}
