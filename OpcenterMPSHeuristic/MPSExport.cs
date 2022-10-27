using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpcenterMPSHeuristic
{
    class MPSExport
    {
        public string ItemCode { get; set; }
        public double CapacityUsed { get; set; }
        public string Period { get; set; }
        public double OnHand { get; set; }//
        public double InitialInventory { get; set; }
        public double GrossRequirements { get; set; }//
        public double StandardLotSize { get; set; }//
        public double NetRequirements { get; set; }//
        public double MPS { get; set; }
        public double TargetStock { get; set; }
        public double MinStock { get; set; }
        public double ClosingStock { get; set; }
        public double MinDaysCover { get; set; }
        public double TargetDaysCover { get; set; }
        public double TotalDaysCover { get; set; }
        public string PlanningResource { get; set; }
        public int RequirementsMet { get; set; }//
        public int RequirementsNotMet { get; set; }//
        public int ServiceLevel { get; set; }//
        public double AverageServiceLevelPeriod { get; set; }//
        public double EndingInventory { get; set; }//
        public double AverageInventoryPeriod { get; set; }//
        public double TotalAverageInventoryPeriod { get; set; }//
        public double BelowSafetyInvetory { get; set; }//
        public double BelowSafetyInvetoryPeriod { get; set; }//
    }
}
