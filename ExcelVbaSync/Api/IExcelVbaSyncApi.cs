using ExcelVbaSync.Sync;
using ExcelVbaSync.Sync.Export;
using ExcelVbaSync.Sync.Import;
using Microsoft.Office.Interop.Excel;

namespace ExcelVbaSync.Api
{
    /// <summary>
    /// Main entrypoint to the library.
    /// </summary>
    public interface IExcelVbaSyncApi
    {
        /// <summary>
        /// Build a new instance of <see cref="IExcelVbaExporter"/> for the provided workbook.
        /// </summary>
        IExcelVbaExporter NewVbaExporter(Workbook workbook);

        /// <summary>
        /// Build a new instance of <see cref="IExcelVbaImporter"/> for the provided workbook.
        /// </summary>
        IExcelVbaImporter NewVbaImporter(Workbook workbook);
    }
}
