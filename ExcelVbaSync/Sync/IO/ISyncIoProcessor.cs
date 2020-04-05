using ExcelVbaSync.Vba;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelVbaSync.Sync.IO
{
    interface ISyncIoProcessor
    {
        void RemoveEmptyLinesFromEndOfFile(string filePath);

        string GetComponentExportName(IVbComponentDecorator component);
    }
}
