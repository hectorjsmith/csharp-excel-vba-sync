using ExcelVbaSync.Vba;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelVbaSync.Sync.IO
{
    interface ISyncFileProcessor
    {
        void AssertPathIsDirectory(string directoryPath);

        void RemoveEmptyLinesFromEndOfFile(string filePath);

        string GetComponentNameFromFileName(string fileName);

        string GetComponentExportName(IVbComponentDecorator component);
    }
}
