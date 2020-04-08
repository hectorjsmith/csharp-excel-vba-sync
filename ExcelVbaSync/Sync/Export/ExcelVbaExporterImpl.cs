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

        private readonly ISyncFileProcessor syncFileProcessor = new SyncFileProcessorImpl();
        private readonly IVbComponentIo vbComponentIo = new VbComponentIoImpl();
        private readonly IVbComponentDecoratorFactory componentFactory = new VbComponentDecoratorFactoryImpl();

        public ExcelVbaExporterImpl(Workbook workbook)
        {
            _workbook = workbook;
        }

        public void Export(string outputDirectory, Func<IVbComponentDecorator, bool> vbComponentFilter)
        {
            syncFileProcessor.AssertPathIsDirectory(outputDirectory);

            IEnumerable<IVbComponentDecorator> filteredComponents = componentFactory
                .GetDecoratedComponentsFromWorkbook(_workbook)
                .Where(vbComponentFilter);

            foreach (IVbComponentDecorator component in filteredComponents)
            {
                string fileName = syncFileProcessor.GetComponentExportName(component);
                string fullPath = Path.Combine(outputDirectory, fileName);

                vbComponentIo.ExportCodeToFile(component, fullPath);
                syncFileProcessor.RemoveEmptyLinesFromEndOfFile(fullPath);
            }
        }

        public void Export(string outputDirectory)
        {
            Export(outputDirectory, component => true);
        }
    }
}
