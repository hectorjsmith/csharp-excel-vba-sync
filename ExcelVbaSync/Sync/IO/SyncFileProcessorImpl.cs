using ExcelVbaSync.Vba;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelVbaSync.Sync.IO
{
    class SyncFileProcessorImpl : ISyncFileProcessor
    {
        private const string ThisWorkbookComponentName = "ThisWorkbook";
        private const string SheetNameSeparatorString = " - ";

        public void AssertPathIsDirectory(string directoryPath)
        {
            FileAttributes attr = File.GetAttributes(directoryPath);
            if (!attr.HasFlag(FileAttributes.Directory))
            {
                throw new InvalidOperationException("Invalid path, must be a directory: " + directoryPath);
            }
        }

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
            VbComponentType componentType = component.ComponentType;
            string componentName = component.Name;
            if (componentType == VbComponentType.Sheet && componentName != ThisWorkbookComponentName)
            {
                return componentName + SheetNameSeparatorString + component.PrettyName + componentType.FileExt;
            }
            else
            {
                return componentName + componentType.FileExt;
            }
        }

        public string GetComponentNameFromFileName(string fileName)
        {
            string fileExt = Path.GetExtension(fileName);
            string rawFilename = Path.GetFileNameWithoutExtension(fileName);

            if (fileExt == VbComponentType.Sheet.FileExt &&
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
