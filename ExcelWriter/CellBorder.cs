using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelWriterCSharp
{
    public class CellBorder
    {
        public Color Color { get; set; }
        public int Thickness { get; set; }
        public BorderStyle Style { get; set; }
        public BorderPosition Position { get; set; }

        public CellBorder(BorderPosition position)
        {
            this.Position = position;
        }
    }

    public enum BorderStyle
    {
        Continuous = Excel.XlLineStyle.xlContinuous,
        Dash = Excel.XlLineStyle.xlDash,
        DashDot = Excel.XlLineStyle.xlDashDot,
        DashDotDot = Excel.XlLineStyle.xlDashDotDot,
        Dot = Excel.XlLineStyle.xlDot,
        Double = Excel.XlLineStyle.xlDouble,
        None = Excel.XlLineStyle.xlLineStyleNone,
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
