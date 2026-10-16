using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPOPComparisonTool
{
    internal class Bullet
    {
        public Bullet() { }

        public string ExcelSPOPName { get; set; }

        public int Row { get; set; }

        public string BulletPoint { get; set; }

        public List<(string Requirement, int Row)> ManualRequirement { get; set; }

        public bool IsNavigateToAnotherSPOP { get; set; } = false;

        public string NavigateToSPOPName { get; set; }

        public bool HasSPSinManualRequirement { get; set; } = false;

        public string SPSinManualRequirement { get; set; }

        public string ProcedureText { get; set; }

        public string SPOPName { get; set; }
    }
}
