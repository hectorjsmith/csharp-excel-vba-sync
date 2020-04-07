using ExcelVbaSync.Vba;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelVbaSync.Sync.IO
{
    interface IVbComponentIo
    {
        void DeleteAllCodeFromComponent(IVbComponentDecorator component);

        void ImportComponentCodeFromFile(IVbComponentDecorator component, string inputFilePath);

        void ImportComponentFromFile(Workbook workbook, string componentFilePath);

        void DeleteComponentFromWorkbook(Workbook workbook, IVbComponentDecorator component);

        void ExportCodeToFile(IVbComponentDecorator component, string filePath);

        string GetVbCodeLines(IVbComponentDecorator component, int numberOfLines);

        int CountCodeLines(IVbComponentDecorator component);

        void DeleteVbCodeLines(IVbComponentDecorator component, int numberOfLines);
    }
}
