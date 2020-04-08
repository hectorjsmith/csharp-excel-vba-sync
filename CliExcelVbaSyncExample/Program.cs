using CommandLine;
using ExcelVbaSync.Api;
using ExcelVbaSync.Sync.Export;
using ExcelVbaSync.Sync.Import;
using Microsoft.Office.Interop.Excel;
using System;

namespace CliExcelVbaSyncExample
{
    class Program
    {
        static int Main(string[] args)
        {
            return Parser.Default.ParseArguments<ImportOptions, ExportOptions>(args)
               .MapResult(
                 (ImportOptions opts) => RunImportAndReturnExitCode(opts),
                 (ExportOptions opts) => RunExportAndReturnExitCode(opts),
                 errs => 1);
        }

        private static void RunAgainstWorkbook(string workbookPath, System.Action<Workbook> actionToRun, bool saveWorkbook)
        {
            Application app = new Application();
            Workbook workbook = app.Workbooks.Open(workbookPath);

            actionToRun(workbook);

            if (saveWorkbook)
            {
                workbook.Save();
            }

            workbook.Close(false);
            app.Quit();
        }

        private static int RunExportAndReturnExitCode(ExportOptions opts)
        {
            RunAgainstWorkbook(
                opts.SourceWorkbookPath,
                wb => {
                    IExcelVbaExporter exporter = ExcelVbaSyncApi.Instance.NewVbaExporter(wb);
                    exporter.Export(opts.TargetDirectoryPath);
                },
                false);

            return 0;
        }

        private static int RunImportAndReturnExitCode(ImportOptions opts)
        {
            RunAgainstWorkbook(
                opts.TargetWorkbookPath,
                wb => {
                    IExcelVbaImporter importer = ExcelVbaSyncApi.Instance.NewVbaImporter(wb);
                    importer.Import(opts.SourceDirectoryPath);
                    importer.RemoveComponentsThatWereNotImported();
                },
                true);

            return 0;
        }
    }
}
