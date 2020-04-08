using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelVbaSync.Sync.Import
{
    public interface IExcelVbaImporter
    {
        void Import(string inputDirectory, Func<string, bool> fileNameFilter);

        void Import(string inputDirectory);

        void RemoveComponentsThatWereNotImported();
    }
}
