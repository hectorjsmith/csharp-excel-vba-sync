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

        private readonly Lazy<ISyncIoProcessor> syncFileProcessor = new Lazy<ISyncIoProcessor>(() => new SyncIoProcessorImpl());
        private readonly Lazy<ISyncComponentIo> syncComponentIo = new Lazy<ISyncComponentIo>(() => new SyncComponentIoImpl());
        private readonly Lazy<IVbComponentDecoratorFactory> componentFactory = new Lazy<IVbComponentDecoratorFactory>(() => new VbComponentDecoratorFactoryImpl());

        public ExcelVbaExporterImpl(Workbook workbook)
        {
            _workbook = workbook;
        }

        public void Export(string outputDirectory, Func<IVbComponentDecorator, bool> vbComponentFilter)
        {
            IEnumerable<IVbComponentDecorator> filteredComponents = componentFactory.Value
                .GetDecoratedComponentsFromWorkbook(_workbook)
                .Where(vbComponentFilter);

            foreach (IVbComponentDecorator component in filteredComponents)
            {
                string fileName = syncFileProcessor.Value.GetComponentExportName(component);
                string fullPath = Path.Combine(outputDirectory, fileName);

                syncComponentIo.Value.ExportCodeToFile(component, fullPath);
                syncFileProcessor.Value.RemoveEmptyLinesFromEndOfFile(fullPath);
            }
        }

        public void Export(string outputDirectory)
        {
            Export(outputDirectory, component => true);
        }
    }
}
