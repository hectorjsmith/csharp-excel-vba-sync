using ExcelVbaSync.Sync.IO;
using ExcelVbaSync.Vba;
using ExcelVbaSync.Vba.Factory;
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

        private readonly Workbook _workbook;
        private readonly string _outputDirectory;

        private readonly Lazy<ISyncIoProcessor> syncFileProcessor = new Lazy<ISyncIoProcessor>(() => new SyncIoProcessorImpl());
        private readonly Lazy<IVbComponentDecoratorFactory> vbComponentFactory = new Lazy<IVbComponentDecoratorFactory>(() => new VbComponentDecoratorFactoryImpl());

        public ExcelVbaExporterImpl(Workbook workbook, string outputDirectory)
        {
            _workbook = workbook;
            _outputDirectory = outputDirectory;
        }

        public void Export(Func<IVbComponentDecorator, bool> vbComponentFilter)
        {
            IEnumerable<IVbComponentDecorator> filteredComponents = vbComponentFactory.Value
                .GetDecoratedComponentsFromWorkbook(_workbook)
                .Where(vbComponentFilter);

            foreach (IVbComponentDecorator component in filteredComponents)
            {
                string fileName = syncFileProcessor.Value.GetComponentExportName(component);
                string fullPath = Path.Combine(_outputDirectory, fileName);

                component.ExportCodeToFile(fullPath);
                syncFileProcessor.Value.RemoveEmptyLinesFromEndOfFile(fullPath);
            }
        }

        public void Export()
        {
            Export(component => true);
        }
    }
}
