using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWriterCSharp
{
	public class TableOptions
	{
        public string Name { get; set; }
        public TableHeaders Headers { get; set; }
        public TableStyle Style { get; set; }

        public TableOptions(string name, TableHeaders headers)
        {
            this.Name = name;
            this.Headers = headers;
        }
	}
}
