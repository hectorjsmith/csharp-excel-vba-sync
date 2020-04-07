using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelVbaSync.Vba
{
    public class VbComponentType
    {
        // Enum Instances
        public static VbComponentType ClassModule { get; } = new VbComponentType("vbext_ct_ClassModule", ".cls");
        public static VbComponentType StandardModule { get; } = new VbComponentType("vbext_ct_StdModule", ".bas");
        public static VbComponentType Sheet { get; } = new VbComponentType("vbext_ct_Document", ".sht");
        public static VbComponentType UserForm { get; } = new VbComponentType("vbext_ct_MSForm", ".frm");

        // Set of Enum values
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
