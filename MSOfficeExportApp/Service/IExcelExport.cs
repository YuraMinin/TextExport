using System;


namespace MSOfficeExportApp.Service
{
    public interface IExcelExport
    {
        void Export(String dir, String text);
    }
}
