using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPOPComparisonTool
{
    internal class Highlight
    {
        public Highlight() { }

        public string SheetName { get; set; }
        public int Row { get; set; }

        public List<int> Columns { get; set; }
    }
}
