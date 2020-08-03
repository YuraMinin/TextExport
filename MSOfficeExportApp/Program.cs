using MSOfficeExportApp.Service;
using MSOfficeExportApp.Service.Impl;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace MSOfficeExportApp
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            MainForm mainForm = new MainForm();
            IWordExport wordExport = new WordExport();
            IExcelExport excelExport = new ExcelExport();

            // Регистрация делегатов
            MainForm.Export exportText = wordExport.Export;
            exportText += excelExport.Export;
            mainForm.registerExportDelegate(exportText);

            Application.Run(mainForm);
        }
    }
}
