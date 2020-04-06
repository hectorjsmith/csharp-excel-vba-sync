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

        private readonly ISet<string> moduleNamesImported = new HashSet<string>();
        private readonly Lazy<ISyncIoProcessor> syncFileProcessor = new Lazy<ISyncIoProcessor>(() => new SyncIoProcessorImpl());
        private readonly Lazy<IVbComponentDecoratorFactory> vbComponentFactory = new Lazy<IVbComponentDecoratorFactory>(() => new VbComponentDecoratorFactoryImpl());

        public ExcelVbaImporterImpl(Workbook workbook)
        {
            _workbook = workbook;
        }

        public void Import(string inputDirectory, Func<string, bool> fileNameFilter)
        {
            ISet<string> pathsToImport = FilePathsToImport(inputDirectory, fileNameFilter);

            foreach (string filePath in pathsToImport)
            {
                string componentName = syncFileProcessor.Value.GetModuleNameFromFileName(filePath);
                IVbComponentDecorator? component = vbComponentFactory.Value
                        .GetVbComponentDecoratorByName(_workbook, componentName);

                ImportVbComponent(component, filePath);
                moduleNamesImported.Add(componentName);
            }
        }

        public void Import(string inputDirectory)
        {
            Import(inputDirectory, str => true);
        }

        public void RemoveComponentsThatWereNotImported()
        {
            RemoveModulesThatNoLongerExist(moduleNamesImported);
        }

        private void ImportVbComponent(IVbComponentDecorator? component, string filePath)
        {
            if (component == null)
            {
                PlainImportComponent(filePath);
            }
            else
            {
                if (component.ComponentType == VbComponentType.VBCompTypeForm)
                {
                    DeleteThenImportModule(component, filePath);
                }
                else
                {
                    ClearThenWriteImportModule(component, filePath);
                }
            }
        }

        private ISet<string> FilePathsToImport(string directoryPath, Func<string, bool> fileNameFilter)
        {
            return Directory.GetFiles(directoryPath)
                .Where(path => fileNameFilter(path) && VbComponentType.Values.Any(type => type.FileExt == Path.GetExtension(path)))
                .ToHashSet();
        }

        private void DeleteThenImportModule(IVbComponentDecorator component, string importFile)
        {
            DeleteComponentFromWorkbook(component);
            ImportComponentFromFile(importFile);
        }

        private void ClearThenWriteImportModule(IVbComponentDecorator component, string importFile)
        {
            // Delete all lines in existing code
            int lineCount = component.CountCodeLines();
            component.DeleteAllCode();

            // Load new code
            component.ImportCodeFromFile(importFile);

            // Cleanup module after import
            CleanupModuleAfterImport(component);
        }

        private void PlainImportComponent(string componentFilePath)
        {
            string fileExt = Path.GetExtension(componentFilePath);
            VbComponentType componentType = VbComponentType.Values
                .First(type => type.FileExt.Equals(fileExt, StringComparison.OrdinalIgnoreCase));

            if (componentType == VbComponentType.VBCompTypeDocument)
            {
                //Log.Warn(string.Format("Using incorrect import method for file: '{0}' - module type '{1}'", importFile, vbCompType.VbCompTypeCode));
                return;
            }
            ImportComponentFromFile(componentFilePath);
        }

        private void RemoveModulesThatNoLongerExist(ISet<string> importModuleNames)
        {
            foreach (IVbComponentDecorator vbComp in vbComponentFactory.Value.GetDecoratedComponentsFromWorkbook(_workbook))
            {
                // If the module exists in the tool, but was not in the import file list, remove it
                // This assumes that the import process always imports everything
                if (!importModuleNames.Contains(vbComp.Name))
                {
                    //Log.Info(string.Format("Module was not part of import set and was deleted: '{0}'", vbComp.GetComponentRawName()));
                    DeleteComponentFromWorkbook(vbComp);
                }
            }
        }

        private void ImportComponentFromFile(string componentFilePath)
        {
            _workbook.VBProject.VBComponents.Import(componentFilePath);
        }

        private void DeleteComponentFromWorkbook(IVbComponentDecorator component)
        {
            _workbook.VBProject.VBComponents.Remove(component.RawComponent);
        }

        private void CleanupModuleAfterImport(IVbComponentDecorator component)
        {
            VbComponentType vbCompType = component.ComponentType;
            string headerText;
            if (vbCompType == VbComponentType.VBCompTypeClassModule || vbCompType == VbComponentType.VBCompTypeDocument)
            {
                // Delete header lines in sheets and classes
                headerText = component.GetVbCodeLines(4);
                if (headerText.ToLower() == "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEnd".ToLower())
                {
                    component.DeleteVbCodeLines(4);
                }
            }
            if (vbCompType == VbComponentType.VBCompTypeForm)
            {
                // Delete header lines in forms
                headerText = component.GetVbCodeLines(10);
                if (headerText.StartsWith("version 5#\r\nbegin {", StringComparison.OrdinalIgnoreCase) &&
                        headerText.EndsWith("end", StringComparison.OrdinalIgnoreCase))
                {
                    component.DeleteVbCodeLines(10);
                }
            }
        }

    }
}
