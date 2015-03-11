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
}
