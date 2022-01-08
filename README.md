# C# Excel VBA Sync

[![project version on nuget](https://badgen.net/nuget/v/ExcelVbaSync/latest)](https://www.nuget.org/packages/ExcelVbaSync/)

C# library to support importing and exporting VBA code from an Excel workbook

## How to Use

The main entry point to the library is the `ExcelVbaSyncApi` singleton class. From that class you can get a VBA importer or exporter for an Excel Workbook object (`Microsoft.Office.Interop.Excel.Workbook`).

There is an example CLI project included in this repository to show how to use this library.

### Export

The exporter class allows exporting VBA code from a given workbook to a folder on you local machine.

```c#
IExcelVbaExporter exporter = ExcelVbaSyncApi.Instance.NewVbaExporter(workbook);
exporter.Export(targetDirectoryPath);
```

You can also filter what VBA components you want to export by providing a filter function.

```c#
exporter.Export(targetDirectoryPath, component => component.Name.Contains("Export"));
```

The exporter class can be re-used multiple times to export the VBA code to different folders or use different filters.

```c#
exporter.Export(firstPath, component => component.Name.StartsWith("A_"));
exporter.Export(secondPath, component => component.Name.StartsWith("B_"));
```

### Import

The import functionality will import VBA code into the given workbook and replace the code that is currently there. This expects the VBA code to be in the same format that the export functionality generates.

```c#
IExcelVbaImporter importer = ExcelVbaSyncApi.Instance.NewVbaImporter(workbook);
importer.Import(sourceDirectoryPath);
importer.RemoveComponentsThatWereNotImported();
```

Similar to the export functionality it is possible to provide a filter to choose which files get imported.

```c#
importer.Import(sourceDirectoryPath, fileName => fileName.Contains("Import"));
```

The importer can also be used to import VBA code multiple times from different folders or using different filters.

```c#
importer.Import(firstPath, fileName => fileName.StartsWith("A_"));
importer.Import(secondPath, fileName => fileName.StartsWith("B_"));
```

#### RemoveComponentsThatWereNotImported

The importer has an extra function: `RemoveComponentsThatWereNotImported()`. The purpose of this function is to clean up the existing code in the Workbook by removing VBA components that were not imported. The purpose is to have the VBA code in the workbook mirror the VBA code in the folder it was imported from.

For example, if the Workbook contains the following components: `A`, `B`, and `C`, but only the `A` and `B` components were imported, this function will remove component `C`.

This makes sense when using this library to version control VBA code in a workbook, because after an import you will want the VBA code in the workbook to always mirror the VBA code in the source control folder.

