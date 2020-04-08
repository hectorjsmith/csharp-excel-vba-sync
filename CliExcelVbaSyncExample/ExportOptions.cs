using CommandLine;
using System;
using System.Collections.Generic;
using System.Text;

namespace CliExcelVbaSyncExample
{
    [Verb("export", HelpText = "Export VBA code from an Excel Workbook file to a directory")]
    class ExportOptions
    {
        [Option('s', "source", Required = true, HelpText = "Path to the Excel Workbook to export the VBA code from")]
        public string SourceWorkbookPath { get; set; } = "";

        [Option('t', "target", Required = true, HelpText = "Path to the directory to write the VBA code files to")]
        public string TargetDirectoryPath { get; set; } = "";
    }
}
