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
    /// <summary>
    /// Main entrypoint to the library. Use the <see cref="Instance"/> property to get an actual instance of this type.
    /// </summary>
    public class ExcelVbaSyncApi : IExcelVbaSyncApi
    {
        /// <summary>
        /// Singleton instance of the sync API.
        /// </summary>
        public static IExcelVbaSyncApi Instance { get; } = new ExcelVbaSyncApi();

        private ExcelVbaSyncApi()
        {
        }

        /// <summary>
        /// Build a new instance of <see cref="IExcelVbaExporter"/> for the provided workbook.
        /// </summary>
        public IExcelVbaExporter NewVbaExporter(Workbook workbook)
        {
            return new ExcelVbaExporterImpl(workbook);
        }

        /// <summary>
        /// Build a new instance of <see cref="IExcelVbaImporter"/> for the provided workbook.
        /// </summary>
        public IExcelVbaImporter NewVbaImporter(Workbook workbook)
        {
            return new ExcelVbaImporterImpl(workbook);
        }
    }
}
