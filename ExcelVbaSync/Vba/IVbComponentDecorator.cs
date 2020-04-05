using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelVbaSync.Vba
{
    public interface IVbComponentDecorator
    {
        string Name { get; }

        string PrettyName { get; }

        VBComponent RawComponent { get; }

        VbComponentType ComponentType { get; }

        string GetVbCodeLines(int numberOfLines);

        void DeleteVbCodeLines(int numberOfLines);

        void DeleteAllCode();

        int CountCodeLines();

        void ImportCodeFromFile(string filePath);

        void ExportCodeToFile(string filePath);
    }
}
