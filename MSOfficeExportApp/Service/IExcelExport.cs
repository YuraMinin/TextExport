using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSOfficeExportApp.Service
{
    public interface IExcelExport
    {
        void Export(String dir, String text);
    }
}
