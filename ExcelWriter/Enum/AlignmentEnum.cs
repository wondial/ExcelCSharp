using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelWriterCSharp
{
    public enum HorizontalAlignment
    {
        None,
        Center = Excel.XlHAlign.xlHAlignCenter,
        Left = Excel.XlHAlign.xlHAlignLeft,
        Right = Excel.XlHAlign.xlHAlignRight,
        Justify = Excel.XlHAlign.xlHAlignJustify
    }

    public enum VerticalAlignment
    {
        None,
        Center = Excel.XlVAlign.xlVAlignCenter,
        Top = Excel.XlVAlign.xlVAlignTop,
        Bottom = Excel.XlVAlign.xlVAlignBottom,
        Justify = Excel.XlVAlign.xlVAlignJustify
    }
}
