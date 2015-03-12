using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;

namespace ExcelWriterCSharp
{
    public class CellOptions
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }

        public string FontName { get; set; }
        public int FontSize { get; set; }
        public Color FontColor { get; set; }

        public Color CellColor { get; set; }

        public HorizontalAlignment TextHorizontalAlignment { get; set; }
        public VerticalAlignment TextVerticalAlignment { get; set; }

        public ReadOnlyCollection<CellBorder> Borders
        {
            get { return borders.AsReadOnly(); }
        } 

        private List<CellBorder> borders;

        public CellOptions()
        {
            borders = new List<CellBorder>();
        }

        public void AddBorder(CellBorder border)
        {
            borders.Add(border);
        }

        public void RemoveBorder(CellBorder border)
        {
            borders.Remove(border);
        }
    }
}
