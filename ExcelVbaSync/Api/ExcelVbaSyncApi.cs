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

        public IExcelVbaExporter NewVbaExporter(Workbook workbook, string outputDirectory)
        {
            AssertPathIsDirectory(outputDirectory);
            return new ExcelVbaExporterImpl(workbook, outputDirectory);
        }

        public IExcelVbaImporter NewVbaImporter(Workbook workbook, string inputDirectory)
        {
            AssertPathIsDirectory(inputDirectory);
            return new ExcelVbaImporterImpl(workbook, inputDirectory);
        }

        private void AssertPathIsDirectory(string path)
        {
            FileAttributes attr = File.GetAttributes(path);
            if (!attr.HasFlag(FileAttributes.Directory))
            {
                throw new InvalidOperationException("Invalid path, must be a directory: " + path);
            }
        }
    }
}
