using ExcelVbaSync.Sync;
using ExcelVbaSync.Sync.Export;
using ExcelVbaSync.Sync.Import;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelVbaSync.Api
{
    public class ExcelVbaSyncApi : IExcelVbaSyncApi
    {
        public static IExcelVbaSyncApi Instance { get; } = new ExcelVbaSyncApi();

        private ExcelVbaSyncApi()
        {
        }

        public IExcelVbaExporter NewVbaExporter(Workbook workbook)
        {
            return new ExcelVbaExporterImpl(workbook);
        }

        public IExcelVbaImporter NewVbaImporter(Workbook workbook)
        {
            return new ExcelVbaImporterImpl(workbook);
        }
    }
}
