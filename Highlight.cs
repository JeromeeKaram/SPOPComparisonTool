using System.Collections.Generic;

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
