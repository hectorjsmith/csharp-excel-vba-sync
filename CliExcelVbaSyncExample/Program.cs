using ExcelVbaSync.Api;
using ExcelVbaSync.Sync.Export;
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

            workbook.Close(false);
            app.Quit();
        }
    }
}
