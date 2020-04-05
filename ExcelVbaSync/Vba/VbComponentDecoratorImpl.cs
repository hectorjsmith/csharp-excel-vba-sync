using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelVbaSync.Vba
{
    class VbComponentDecoratorImpl : IVbComponentDecorator
    {
        private const int SheetNamePropertyIndex = 7;

        public VBComponent RawComponent { get; }

        public VbComponentType ComponentType { get; }

        public string Name => RawComponent.Name;

        public string PrettyName => RawComponent.Properties.Item(SheetNamePropertyIndex).Value.ToString() ?? string.Empty;

        public VbComponentDecoratorImpl(VBComponent rawComponent, VbComponentType componentType)
        {
            RawComponent = rawComponent;
            ComponentType = componentType;
        }

        public string GetVbCodeLines(int numberOfLines)
        {
            return RawComponent.CodeModule.Lines[1, numberOfLines];
        }

        public void DeleteVbCodeLines(int numberOfLines)
        {
            RawComponent.CodeModule.DeleteLines(1, numberOfLines);
        }

        public void DeleteAllCode()
        {
            DeleteVbCodeLines(CountCodeLines());
        }

        public int CountCodeLines()
        {
            return RawComponent.CodeModule.CountOfLines;
        }

        public void ImportCodeFromFile(string filePath)
        {
            RawComponent.CodeModule.AddFromFile(filePath);
        }

        public void ExportCodeToFile(string filePath)
        {
            RawComponent.Export(filePath);
        }
    }
}
