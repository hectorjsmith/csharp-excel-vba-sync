using ExcelVbaSync.Vba;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelVbaSync.Sync.IO
{
    class SyncIoProcessorImpl : ISyncIoProcessor
    {
        private const string ThisWorkbookModuleName = "ThisWorkbook";
        private const string SheetNameSeparatorString = " - ";

        public void RemoveEmptyLinesFromEndOfFile(string filePath)
        {
            string[] lines = File.ReadAllLines(filePath);
            int maxIndex = lines.Count();

            for (int i = lines.Count() - 1; i > 0; i--)
            {
                if (!string.IsNullOrWhiteSpace(lines[i]))
                {
                    break;
                }
                maxIndex--;
            }

            File.WriteAllLines(filePath, lines.Take(maxIndex));
        }

        public string GetComponentExportName(IVbComponentDecorator component)
        {
            VbComponentType vbCompType = component.ComponentType;
            string compName = component.Name;
            if (vbCompType == VbComponentType.VBCompTypeDocument && compName != ThisWorkbookModuleName)
            {
                return compName + SheetNameSeparatorString + component.PrettyName + vbCompType.FileExt;
            }
            else
            {
                return compName + vbCompType.FileExt;
            }
        }

        public string GetModuleNameFromFileName(string fileName)
        {
            string fileExt = Path.GetExtension(fileName);
            string rawFilename = Path.GetFileNameWithoutExtension(fileName);

            if (fileExt == VbComponentType.VBCompTypeDocument.FileExt &&
                    rawFilename.Contains(SheetNameSeparatorString))
            {
                // Get substring up to first dash for sheet files
                return rawFilename.Substring(0, rawFilename.IndexOf(SheetNameSeparatorString));
            }
            else
            {
                // No dashes in file name, return plain file name
                return rawFilename;
            }
        }
    }
}
