using ExcelVbaSync.Vba;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelVbaSync.Sync.IO
{
    class SyncComponentIoImpl : ISyncComponentIo
    {
        public void DeleteAllCodeFromComponent(IVbComponentDecorator component)
        {
            DeleteVbCodeLines(component, CountCodeLines(component));
        }

        public void DeleteComponentFromWorkbook(Workbook workbook, IVbComponentDecorator component)
        {
            workbook.VBProject.VBComponents.Remove(component.RawComponent);
        }

        public void ImportComponentCodeFromFile(IVbComponentDecorator component, string inputFilePath)
        {
            component.RawComponent.CodeModule.AddFromFile(inputFilePath);
        }

        public void ImportComponentFromFile(Workbook workbook, string componentFilePath)
        {
            workbook.VBProject.VBComponents.Import(componentFilePath);
        }

        public string GetVbCodeLines(IVbComponentDecorator component, int numberOfLines)
        {
            return component.RawComponent.CodeModule.Lines[1, numberOfLines];
        }

        public void DeleteVbCodeLines(IVbComponentDecorator component, int numberOfLines)
        {
            component.RawComponent.CodeModule.DeleteLines(1, numberOfLines);
        }

        public int CountCodeLines(IVbComponentDecorator component)
        {
            return component.RawComponent.CodeModule.CountOfLines;
        }

        public void ExportCodeToFile(IVbComponentDecorator component, string filePath)
        {
            component.RawComponent.Export(filePath);
        }
    }
}
