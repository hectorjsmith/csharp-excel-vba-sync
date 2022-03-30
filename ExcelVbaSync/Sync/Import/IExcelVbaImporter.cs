using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelVbaSync.Sync.Import
{
    /// <summary>
    /// Type responsible for importing VBA code from the filesystem into a workbook.
    /// </summary>
    public interface IExcelVbaImporter
    {
        /// <summary>
        /// Import VBA code from the provided directory into the workbook. Each file found will be passed through the provided filter and only files where the filter function returs true will be imported.
        /// </summary>
        /// <param name="inputDirectory">Directory to read VBA code files from.</param>
        /// <param name="fileNameFilter">Filter function used to determine which files to import.</param>
        void Import(string inputDirectory, Func<string, bool> fileNameFilter);

        /// <summary>
        /// Import VBA code from the provided directory into the workbook.
        /// </summary>
        /// <param name="inputDirectory">Directory to read VBA code files from.</param>
        void Import(string inputDirectory);

        /// <summary>
        /// Remove all components that were not imported by this importer instance.
        /// Each <see cref="IExcelVbaImporter"/> instance keeps track of all the modules it has imported so far, when this method is called, all modules that haven't yet been imported are removed from the workbook.
        /// </summary>
        void RemoveComponentsThatWereNotImported();
    }
}
