using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using System.Collections.Generic;

namespace ExcelVbaSync.Vba.Factory
{
    interface IVbComponentDecoratorFactory
    {
        IEnumerable<IVbComponentDecorator> GetDecoratedComponentsFromWorkbook(Workbook workbook);

        IVbComponentDecorator? GetVbComponentDecoratorByName(Workbook workbook, string componentName);

        IVbComponentDecorator MapVbComponentToVbComponentDecorator(VBComponent rawComponent);
    }
}
