using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpcenterMPSHeuristic
{
    class Item
    {
        public string ItemCode { get; set; }
        public string ItemDesc { get; set; }
        public string ItemLevel { get; set; }
        public string CapacityUoM { get; set; }
        public string PlanningResourceGroup { get; set; }
        public double ReorderMultiple { get; set; }                 //Tamanho de lote padrão
        public double MinimumReorderMultiple { get; set; }          //Tamanho mínimo de lote
        public double MinimumCoverDays { get; set; }
        public double TargetCoverDays { get; set; }
        public double MaximumCoverDays { get; set; }
        public double NetRequirements { get; set; }


    }
}
