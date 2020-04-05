using ExcelVbaSync.Api;
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

            ExcelVbaSyncApi.Instance.NewVbaExporter(workbook, outputDirectory).Export();

            workbook.Close(false);
            app.Quit();
        }
    }
}
