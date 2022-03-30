using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelVbaSync.Vba
{
    /// <summary>
    /// Decoration type that wraps around the raw <see cref="VBComponent"/> type.
    /// </summary>
    public interface IVbComponentDecorator
    {
        /// <summary>
        /// Name of the component.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Prettified name of the component. In the case of <see cref="VbComponentType.Sheet"/> components this represents the name of the sheet where <see cref="Name"/> represents the sheet ID (e.g. Sheet1).
        /// </summary>
        string PrettyName { get; }

        /// <summary>
        /// Raw <see cref="VBComponent"/> instance this class is wrapping around.
        /// </summary>
        VBComponent RawComponent { get; }

        /// <summary>
        /// Type of component this represents.
        /// </summary>
        VbComponentType ComponentType { get; }
    }
}
