using ExcelVbaSync.Vba;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelVbaSync.Sync.Export
{
    /// <summary>
    /// Type responsible for exporting VBA code out of a workbook and into a file system.
    /// </summary>
    public interface IExcelVbaExporter
    {
        /// <summary>
        /// Export all VBA components that match the provided filter into the provided path.
        /// </summary>
        /// <param name="outputDirectory">Path to the folder where all VBA code files will be saved.</param>
        /// <param name="vbComponentFilter">Filter function that is applied to each VBA component. Only components where the filter returns true will be exported.</param>
        void Export(string outputDirectory, Func<IVbComponentDecorator, bool> vbComponentFilter);

        /// <summary>
        /// Export all VBA code from the workbook into the provided path.
        /// </summary>
        /// <param name="outputDirectory">Path to the folder where all VBA code files will be saved.</param>
        void Export(string outputDirectory);
    }
}
