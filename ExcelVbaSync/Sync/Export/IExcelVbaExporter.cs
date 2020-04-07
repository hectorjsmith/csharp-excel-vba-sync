using ExcelVbaSync.Vba;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelVbaSync.Sync.Export
{
    public interface IExcelVbaExporter
    {
        void Export(string outputDirectory, Func<IVbComponentDecorator, bool> vbComponentFilter);

        void Export(string outputDirectory);
    }
}
