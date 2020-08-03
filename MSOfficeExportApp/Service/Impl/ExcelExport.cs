using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Authentication.ExtendedProtection.Configuration;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace MSOfficeExportApp.Service.Impl
{
    public class ExcelExport : IExcelExport
    {
        public void Export(String dir, String text)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;

            // Создание рабочей книги
            Excel.Workbook workbook = excel.Workbooks.Add();

            // Создание листа
            Excel.Worksheet worksheet = (Excel.Worksheet)excel.Worksheets.get_Item(1);
            worksheet.Name = "Example";

            // Добавление в первую ячейку текста
            worksheet.Cells[1, 1] = text;

            // Сохранение таблицы
            workbook.SaveAs(dir);
            workbook.Close();
            excel.Quit();
        }
    }
}
