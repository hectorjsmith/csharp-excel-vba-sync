using ExcelVbaSync.Sync.IO;
using ExcelVbaSync.Vba;
using ExcelVbaSync.Vba.Factory;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelVbaSync.Sync.Import
{
    class ExcelVbaImporterImpl : IExcelVbaImporter
    {
        private readonly Workbook _workbook;

        private readonly ISet<string> componentNamesImported = new HashSet<string>();
        private readonly ISyncFileProcessor syncFileProcessor = new SyncFileProcessorImpl();
        private readonly IVbComponentIo vbComponentIo = new VbComponentIoImpl();
        private readonly IVbComponentDecoratorFactory cmponentFactory = new VbComponentDecoratorFactoryImpl();

        public ExcelVbaImporterImpl(Workbook workbook)
        {
            _workbook = workbook;
        }

        public void Import(string inputDirectory, Func<string, bool> fileNameFilter)
        {
            syncFileProcessor.AssertPathIsDirectory(inputDirectory);

            ISet<string> pathsToImport = FilePathsToImport(inputDirectory, fileNameFilter);

            foreach (string filePath in pathsToImport)
            {
                string componentName = syncFileProcessor.GetComponentNameFromFileName(filePath);
                IVbComponentDecorator? component = cmponentFactory
                    .GetVbComponentDecoratorByName(_workbook, componentName);

                ImportVbComponent(component, filePath);
                componentNamesImported.Add(componentName);
            }
        }

        public void Import(string inputDirectory)
        {
            Import(inputDirectory, str => true);
        }

        public void RemoveComponentsThatWereNotImported()
        {
            RemoveComponentsNotFoundInNameSet(componentNamesImported);
        }

        private void ImportVbComponent(IVbComponentDecorator? component, string filePath)
        {
            if (component == null)
            {
                PlainImportComponent(filePath);
            }
            else
            {
                if (component.ComponentType == VbComponentType.UserForm)
                {
                    DeleteComponentThenImportFresh(component, filePath);
                }
                else
                {
                    ClearAllCodeThenImportAsText(component, filePath);
                }
            }
        }

        private ISet<string> FilePathsToImport(string directoryPath, Func<string, bool> fileNameFilter)
        {
            return Directory.GetFiles(directoryPath)
                .Where(path => fileNameFilter(path) && VbComponentType.Values.Any(type => type.FileExt == Path.GetExtension(path)))
                .ToHashSet();
        }

        private void DeleteComponentThenImportFresh(IVbComponentDecorator component, string importFile)
        {
            vbComponentIo.DeleteComponentFromWorkbook(_workbook, component);
            vbComponentIo.ImportComponentFromFile(_workbook, importFile);
        }

        private void ClearAllCodeThenImportAsText(IVbComponentDecorator component, string importFile)
        {
            // Delete all lines in existing code
            vbComponentIo.DeleteAllCodeFromComponent(component);

            // Load new code
            vbComponentIo.ImportComponentCodeFromFile(component, importFile);

            // Cleanup code after import
            CleanupComponentAfterImport(component);
        }

        private void PlainImportComponent(string componentFilePath)
        {
            string fileExt = Path.GetExtension(componentFilePath);
            VbComponentType componentType = VbComponentType.Values
                .First(type => type.FileExt.Equals(fileExt, StringComparison.OrdinalIgnoreCase));

            if (componentType == VbComponentType.Sheet)
            {
                //Log.Warn(string.Format("Using incorrect import method for file: '{0}' - module type '{1}'", importFile, vbCompType.VbCompTypeCode));
                return;
            }
            vbComponentIo.ImportComponentFromFile(_workbook, componentFilePath);
        }

        private void RemoveComponentsNotFoundInNameSet(ISet<string> componentNameSet)
        {
            foreach (IVbComponentDecorator component in cmponentFactory.GetDecoratedComponentsFromWorkbook(_workbook))
            {
                // If the module exists in the tool, but was not in the import file list, remove it
                // This assumes that the import process always imports everything
                if (!componentNameSet.Contains(component.Name))
                {
                    if (component.ComponentType == VbComponentType.Sheet)
                    {
                        // Not attempting to delete sheet components because Interop does not allow it - removing all code instead
                        vbComponentIo.DeleteAllCodeFromComponent(component);
                    }
                    else
                    {
                        //Log.Info(string.Format("Module was not part of import set and was deleted: '{0}'", vbComp.GetComponentRawName()));
                        vbComponentIo.DeleteComponentFromWorkbook(_workbook, component);
                    }
                }
            }
        }

        private void CleanupComponentAfterImport(IVbComponentDecorator component)
        {
            VbComponentType componentType = component.ComponentType;
            string headerText;
            if (componentType == VbComponentType.ClassModule || componentType == VbComponentType.Sheet)
            {
                // Delete header lines in sheets and classes
                headerText = vbComponentIo.GetVbCodeLines(component, 4);
                if (headerText.ToLower() == "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEnd".ToLower())
                {
                    vbComponentIo.DeleteVbCodeLines(component, 4);
                }
            }
            if (componentType == VbComponentType.UserForm)
            {
                // Delete header lines in forms
                headerText = vbComponentIo.GetVbCodeLines(component, 10);
                if (headerText.StartsWith("version 5#\r\nbegin {", StringComparison.OrdinalIgnoreCase) &&
                        headerText.EndsWith("end", StringComparison.OrdinalIgnoreCase))
                {
                    vbComponentIo.DeleteVbCodeLines(component, 10);
                }
            }
        }

    }
}
