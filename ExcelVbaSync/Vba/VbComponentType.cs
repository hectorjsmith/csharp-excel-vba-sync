using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelVbaSync.Vba
{
    public class VbComponentType
    {
        private static ISet<VbComponentType> _values;

        // Enum Instances
        public static VbComponentType VBCompTypeClassModule { get; } = new VbComponentType("vbext_ct_ClassModule", ".cls");
        public static VbComponentType VBCompTypeStdModule { get; } = new VbComponentType("vbext_ct_StdModule", ".bas");
        public static VbComponentType VBCompTypeDocument { get; } = new VbComponentType("vbext_ct_Document", ".sht");
        public static VbComponentType VBCompTypeForm { get; } = new VbComponentType("vbext_ct_MSForm", ".frm");
        public static ISet<VbComponentType> Values
        {
            get
            {
                if (_values != null)
                {
                    return _values;
                }
                _values = new HashSet<VbComponentType>
                {
                    VBCompTypeClassModule,
                    VBCompTypeStdModule,
                    VBCompTypeDocument,
                    VBCompTypeForm
                };
                return _values;
            }
        }

        // Data stored for each instance
        public string VbCompTypeCode { get; }
        public string FileExt { get; }

        private VbComponentType(string vbCompTypeCode, string fileExt)
        {
            VbCompTypeCode = vbCompTypeCode;
            FileExt = fileExt;
        }
    }
}
