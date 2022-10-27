using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpcenterMPSHeuristic
{
    class Demand
    {
        //public string Origin { get; set; }
        public string ItemCode { get; set; }
        public DateTime OrderDate { get; set; }
        public double Quantity { get; set; }
    }
}
