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

        public string SPOPName { get; set; }
        public int Row { get; set; }

        public List<string> Columns2Highlight { get; set; }
    }
}
