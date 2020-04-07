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
    }
}
