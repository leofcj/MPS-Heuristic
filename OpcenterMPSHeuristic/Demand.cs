using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpcenterMPSHeuristic
{
    class Demand
    {
        public int Number { get; set; }
        public string ItemCode { get; set; }
        public double BeggingStock { get; set; }
        public double NetRequirements { get; set; }
        public DateTime DemandDate { get; set; }
        public double MPS { get; set; }
        public string Resource { get; set; }
        public string Day { get; set; }
        public double Subcontracted { get; set; }
        public double GrossRequirements { get; set; }
        public double MinimumReorderMultiple { get; set; }
        public double ReorderMultiple { get; set; }
        public double MaximumInventoryLevel { get; set; }
        public double CapacityUsed { get; set; }
        public int ResourceCount { get; set; }
        public int ItemFitsInFull { get; set; }
        public int ItemFitsInPartial { get; set; }
        
    }
}
