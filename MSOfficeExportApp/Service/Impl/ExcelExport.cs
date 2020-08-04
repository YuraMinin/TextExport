using System;
using Excel = Microsoft.Office.Interop.Excel;


namespace MSOfficeExportApp.Service.Impl
{
    public class ExcelExport : IExcelExport
    {
        public void Export(String dir, String text)
        {
            Excel._Application excel = new Excel.Application() { Visible = true };
            Excel._Workbook workbook = excel.Workbooks.Add();

            Excel._Worksheet worksheet = (Excel.Worksheet)excel.Worksheets.get_Item(1);
            worksheet.Name = "Example";

            // Добавление в первую ячейку текста
            worksheet.Cells[1, 1] = text;

            // Сохранение таблицы
            workbook.SaveAs(dir + "\\exampleXlsx.xlsx");
            workbook.Close();
            excel.Quit();
        }
    }
}
