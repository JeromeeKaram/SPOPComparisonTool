using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPOPComparisonTool
{
    internal class ExcelData
    {
        public ExcelData() { }

        public string SheetName { get; set; }

        public int Row { get; set; }

        public string BColumnText { get; set; }

        public List<string> GHColumnText { get; set; }
    }
}
