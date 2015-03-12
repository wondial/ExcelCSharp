using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelWriterCSharp
{
    public enum BorderStyle
    {
        None = 0,
        Continuous = Excel.XlLineStyle.xlContinuous,
        Dash = Excel.XlLineStyle.xlDash,
        DashDot = Excel.XlLineStyle.xlDashDot,
        DashDotDot = Excel.XlLineStyle.xlDashDotDot,
        Dot = Excel.XlLineStyle.xlDot,
        Double = Excel.XlLineStyle.xlDouble,
        SlantDashDot = Excel.XlLineStyle.xlSlantDashDot
    }

    public enum BorderPosition
    {
        Top = Excel.XlBordersIndex.xlEdgeTop,
        Left = Excel.XlBordersIndex.xlEdgeLeft,
        Right = Excel.XlBordersIndex.xlEdgeRight,
        Bottom = Excel.XlBordersIndex.xlEdgeBottom
    }
}
