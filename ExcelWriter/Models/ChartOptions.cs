using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWriterCSharp
{
    public class ChartOptions
    {
        public ChartStyle Style { get; set; }
        public string Title { get; set; }
        public double Top { get; set; }
        public double Left { get; set; }
        public string XAxeTitle { get; set; }
        public string YAxeTitle { get; set; }

        public ChartOptions(double left, double top)
        {
            this.Left = left;
            this.Top = top;
        }
    }
}
