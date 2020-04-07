using CommandLine;
using System;
using System.Collections.Generic;
using System.Text;

namespace CliExcelVbaSyncExample
{
    [Verb("import", HelpText = "Import VBA code from a directory into an Excel Workbook file")]
    class ImportOptions
    {
        [Option('s', "source", Required = true, HelpText = "Path to directory that contains exported VBA code files")]
        public string SourceDirectoryPath { get; set; } = "";

        [Option('t', "target", Required = true, HelpText = "Path to the Excel Workbook to import the VBA code into")]
        public string TargetWorkbookPath { get; set; } = "";
    }
}
