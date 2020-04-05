using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelVbaSync.Vba.Factory
{
    class VbComponentDecoratorFactoryImpl : IVbComponentDecoratorFactory
    {
        public IEnumerable<IVbComponentDecorator> GetDecoratedComponentsFromWorkbook(Workbook workbook)
        {
            IVbComponentDecoratorFactory factory = new VbComponentDecoratorFactoryImpl();

            System.Collections.IEnumerable vbComponents = workbook.VBProject.VBComponents;
            return vbComponents
                .Cast<VBComponent>()
                .Select(rawComponent => factory.MapVbComponentToVbComponentDecorator(rawComponent));
        }

        public IVbComponentDecorator MapVbComponentToVbComponentDecorator(VBComponent rawComponent)
        {
            VbComponentType componentType = VbComponentType.Values
                .FirstOrDefault(type => type.VbCompTypeCode == rawComponent.Type.ToString());

            return new VbComponentDecoratorImpl(rawComponent, componentType);
        }
    }
}
