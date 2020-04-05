using ExcelVbaSync.Sync;
using ExcelVbaSync.Sync.Export;
using ExcelVbaSync.Sync.Import;
using Microsoft.Office.Interop.Excel;

namespace ExcelVbaSync.Api
{
    public interface IExcelVbaSyncApi
    {
        IExcelVbaExporter NewVbaExporter(Workbook workbook, string outputDirectory);

        IExcelVbaImporter NewVbaImporter(Workbook workbook, string inputDirectory);
    }
}
