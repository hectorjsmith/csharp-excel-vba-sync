using ExcelVbaSync.Vba;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelVbaSync.Sync.Export
{
    public interface IExcelVbaExporter
    {
        void Export(Func<IVbComponentDecorator, bool> vbComponentFilter);

        void Export();
    }
}
