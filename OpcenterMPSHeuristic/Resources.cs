using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpcenterMPSHeuristic
{
    class Resource
    {

        public string ResourceName { get; set; }
        public double AvailableCapacityPeriodInWeek { get; set; }
        public double AvailableCapacityPeriodInMonth { get; set; }
        public double OriginalCapacityPeriodInWeek { get; set; }
        public double OriginalCapacityPeriodInMonth { get; set; }
        public double OvertimePercent { get; set; }
        public DateTime DatePeriod { get; set; }
        public double CapacityUsed { get; set; }
        public string ItemCode { get; set; }
        public string Day { get; set; }
    }
    
}   
                        