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
        public double AvailableCapacityPeriod { get; set; }
        public DateTime DatePeriod { get; set; }
    }
}
