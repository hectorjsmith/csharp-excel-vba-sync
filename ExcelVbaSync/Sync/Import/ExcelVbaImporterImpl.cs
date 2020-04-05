using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelVbaSync.Sync.Import
{
    class ExcelVbaImporterImpl : IExcelVbaImporter
    {
        private readonly Workbook _workbook;
        private readonly string _inputDirectory;

        public ExcelVbaImporterImpl(Workbook workbook, string inputDirectory)
        {
            _workbook = workbook;
            _inputDirectory = inputDirectory;
        }

        public void Import(Func<string, bool> fileNameFilter)
        {
            throw new NotImplementedException();
        }

        public void Import()
        {
            Import(str => true);
        }
    }
}
