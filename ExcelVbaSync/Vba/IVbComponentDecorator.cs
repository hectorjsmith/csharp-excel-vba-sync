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
    }
}
