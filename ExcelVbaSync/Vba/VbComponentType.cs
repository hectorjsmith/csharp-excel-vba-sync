using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelVbaSync.Vba
{
    /// <summary>
    /// Pseudo-enum type containing all the different types of <see cref="Microsoft.Vbe.Interop.VBComponent"/> objects.
    /// Each instance contains the type code and corresponding file extension.
    /// </summary>
    public class VbComponentType
    {
        // Enum Instances
        /// <summary>
        /// Represents a "Class Module" in VBA.
        /// </summary>
        public static VbComponentType ClassModule { get; } = new VbComponentType("vbext_ct_ClassModule", ".cls");
        /// <summary>
        /// Represents a standard "Module" in VBA.
        /// </summary>
        public static VbComponentType StandardModule { get; } = new VbComponentType("vbext_ct_StdModule", ".bas");
        /// <summary>
        /// Represents a "Microsoft Excel Object" component (i.e. a worksheet component) in VBA.
        /// </summary>
        public static VbComponentType Sheet { get; } = new VbComponentType("vbext_ct_Document", ".sht");
        /// <summary>
        /// Represents a "Form" component in VBA.
        /// </summary>
        public static VbComponentType UserForm { get; } = new VbComponentType("vbext_ct_MSForm", ".frm");

        /// <summary>
        /// Set of all possible component types.
        /// </summary>
        public static ISet<VbComponentType> Values { get; } = CalculateComponentTypeSet();
        private static ISet<VbComponentType> CalculateComponentTypeSet()
        {
            return new HashSet<VbComponentType> {
                ClassModule,
                StandardModule,
                Sheet,
                UserForm
            };
        }

        /// <summary>
        /// Code used by Interop to identify this component type.
        /// </summary>
        public string VbCompTypeCode { get; }

        /// <summary>
        /// File extension to use when saving or reading this component from a file.
        /// </summary>
        public string FileExt { get; }

        private VbComponentType(string vbCompTypeCode, string fileExt)
        {
            VbCompTypeCode = vbCompTypeCode;
            FileExt = fileExt;
        }
    }
}
