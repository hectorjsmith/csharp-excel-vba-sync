using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelVbaSync.Sync.Import
{
    public interface IExcelVbaImporter
    {
        void Import(Func<string, bool> fileNameFilter);

        void Import();
    }
}
