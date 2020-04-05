using ExcelVbaSync.Vba;
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelVbaSync.Sync.Export
{
    class ExcelVbaExporterImpl : IExcelVbaExporter
    {
        private const string ThisWorkbookModuleName = "ThisWorkbook";

        private readonly Workbook _workbook;
        private readonly string _outputDirectory;

        public ExcelVbaExporterImpl(Workbook workbook, string outputDirectory)
        {
            _workbook = workbook;
            _outputDirectory = outputDirectory;
        }

        public void Export(Func<IVbComponentDecorator, bool> vbComponentFilter)
        {
            IEnumerable<IVbComponentDecorator> filteredComponents = GetComponentsFromWorkbook()
                .Select(comp => MapRawVbComponentToDecoratedType(comp))
                .Where(vbComponentFilter);

            foreach (IVbComponentDecorator component in filteredComponents)
            {
                string fileName = GetComponentExportName(component);
                string fullPath = Path.Combine(_outputDirectory, fileName);

                component.ExportCodeToFile(fullPath);
                RemoveEmptyLinesFromEndOfFile(fullPath);
            }
        }

        public void Export()
        {
            Export(component => true);
        }

        private IEnumerable<VBComponent> GetComponentsFromWorkbook()
        {
            System.Collections.IEnumerable vbComponents = _workbook.VBProject.VBComponents;
            return vbComponents.Cast<VBComponent>();
        }

        private IVbComponentDecorator MapRawVbComponentToDecoratedType(VBComponent component)
        {
            VbComponentType componentType = VbComponentType.Values
                .FirstOrDefault(type => type.VbCompTypeCode == component.Type.ToString());

            return new VbComponentDecoratorImpl(component, componentType);
        }

        private string GetComponentExportName(IVbComponentDecorator vbComp)
        {
            VbComponentType vbCompType = vbComp.ComponentType;
            string compName = vbComp.Name;
            if (vbCompType == VbComponentType.VBCompTypeDocument && compName != ThisWorkbookModuleName)
            {
                return compName + " - " + vbComp.PrettyName + vbCompType.FileExt;
            }
            else
            {
                return compName + vbCompType.FileExt;
            }
        }

        private void RemoveEmptyLinesFromEndOfFile(string filePath)
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
    }
}
