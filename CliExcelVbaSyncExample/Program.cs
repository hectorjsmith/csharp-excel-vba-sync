using ExcelVbaSync.Api;
using ExcelVbaSync.Sync.Export;
using ExcelVbaSync.Sync.Import;
using Microsoft.Office.Interop.Excel;
using System;

namespace CliExcelVbaSyncExample
{
    class Program
    {
        static void Main(string[] args)
        {
            string workbookName = args[0];
            string outputDirectory = args[1];

            Application app = new Application();
            Workbook workbook = app.Workbooks.Open(workbookName);

            IExcelVbaExporter exporter = ExcelVbaSyncApi.Instance.NewVbaExporter(workbook);
            exporter.Export(outputDirectory);

            IExcelVbaImporter importer = ExcelVbaSyncApi.Instance.NewVbaImporter(workbook);
            importer.Import(outputDirectory);
            importer.RemoveComponentsThatWereNotImported();

            workbook.Save();
            workbook.Close(false);
            app.Quit();
        }
    }
}
