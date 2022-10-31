using Preactor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



public static class Declarations
    {
    //Tipos nativos da API Opcenter
    public static IPreactor sharedPreactor { get; set; } 


    // Tabelas nativas do Opcenter utilizadas neste desenvolvimento
    public static string tblItem { get; } = "Items";
    public static string tblStock { get; } = "Imported Stocks";
    public static string tblNonAggDemand { get; } = "Non Aggregated Demand";
    public static string tblPlanningResources { get; } = "Planning Resources";
    public static string tblDemand { get; } = "Demand";

    //Campos da tabela de Items
    public static string clnItemsItemCode { get; } = "Item Code";
    public static string clnItemsItemDesc { get; } = "Item Desc";
    public static string clnItemsItemLevel { get; } = "Item Level";
    public static string clnItemsCapacityUoM { get; } = "Capacity UoM";
    public static string clnItemsPlanningResourceGroup { get; } = "Planning Resource Group";
    public static string clnItemsPlanningResourceData { get; } = "Planning Resource Data";

    //public static string clnItemsResourceSpecificRateperHour { get; } = "";
    public static string clnItemsReorderMultiple { get; } = "Reorder Multiple";
    public static string clnItemsMinimumReorderMultiple { get; } = "Minimum Reorder Quantity";
    public static string clnItemsMinimumCoverDays { get; } = "Min Days of Cover";
    public static string clnItemsTargetCoverDays { get; } = "Target Days of Cover";
    public static string clnItemsMaximumCoverDays { get; } = "Absolute Max Days of Cover";
        

    //Campos da tabela de Estoques
    public static string clnStockItemCode { get; } = "Item Code";
    public static string clnStockProdnDate { get; } = "ProdnDate";
    public static string clnStockQuantity { get; } = "Qty";


    //Campos da tabela de Demanda Não Agregada
    public static string clnDemandItemCode { get; } = "Item Code";
    public static string clnDemandOrderDate { get; } = "Order Date";
    public static string clnDemandQuantity { get; } = "Quantity";

    //Campos da tabela de Recursos
    public static string clnResourceName { get; } = "Name";

    //Campos da tabela de Demand
    public static string clnDemandNumber { get; } = "Number";
    public static string clnDemandCode { get; } = "Code";
    public static string clnDemandCapacityUsed { get; } = "Capacity Used";
    public static string clnDemandDate { get; } = "Demand Date";
    public static string clnDemandOpeningStock { get; } = "Opening Stock";
    public static string clnDemandDemand { get; } = "Demand";
    public static string clnDemandTargetStock { get; } = "Target Stock";
    public static string clnDemandMinStock { get; } = "Min Stock";
    public static string clnDemandClosingStock { get; } = "Closing Stock";
    public static string clnDemandMPS { get; } = "MPS";
    public static string clnDemandPlanningResource { get; } = "Planning Resource";
    public static string clnDemandMinDaysCover { get; } = "Min Days of Cover";
    public static string clnDemandTargetDaysCover { get; } = "Target Days of Cover";
    public static string clnDemandTotalDaysCover { get; } = "Total Days of Cover";
    public static string clnDemandSubcontracted { get; } = "Subcontracted";
    public static string clnDemandMinimumLotSize { get; } = "Minimum Reorder Quantity";
    public static string clnDemandReorderMultiple { get; } = "Reorder Multiple";
    public static string clndDemandMaximumInventoryLevel { get; } = "Maximum Inventory Level";
    public static string clnDemandNetRequirements { get; } = "Net Requirements";
    public static string clnDemandSafetyInventory { get; } = "Safety Inventory";

}

