﻿using System;
using System.Runtime.InteropServices;
using Preactor;
using Preactor.Interop.PreactorObject;
using System.Windows.Forms;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using static Declarations;

using Microsoft.Office.Interop.Excel;

namespace OpcenterMPSHeuristic
{
    [Guid("767aaf60-acee-4114-a2ac-3d48db672fc5")]
    [ComVisible(true)]
    public interface IMPSHeuristic
    {

        int genMPS(ref PreactorObj preactorComObject, ref object pespComObject);

    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("8375df0c-6d42-41a6-b03d-08cdbcf8bd7d")]
    public class MPSHeuristic : IMPSHeuristic
    {
        IList<Item> ItemsList = new List<Item>();
        IList<Demand> DemandList = new List<Demand>();
        IList<Stock> StockList = new List<Stock>();
        IList<Resource> ResourceList = new List<Resource>();
        IList<MPSResults> MPSList = new List<MPSResults>();
        IList<MPSExport> MPSExportList = new List<MPSExport>();

        public int genMPS(ref PreactorObj preactorComObject, ref object pespComObject)
        {
            sharedPreactor = PreactorFactory.CreatePreactorObject(preactorComObject);

            //int formatResources = preactor.GetFormatNumber("Planning Resources");
            //int formatResourceNumber = preactor.GetFieldNumber(formatResources, "Name");
            ////int recordnumber = preactor.ReadFieldInt(formatResources, formatResourceNumber, 198);
            //int keyValue = 198;
            //PrimaryKey resourcePrimaryKey = new PrimaryKey(keyValue);

            //int teste = preactor.GetRecordNumber(formatResources, resourcePrimaryKey);
            //string resourceName = preactor.ReadFieldString(formatResources, formatResourceNumber, resourcePrimaryKey);
            //MessageBox.Show(resourceName);

            //int teste = sharedPreactor.FindMatchingRecord(tblStock, clnStockItemCode, 0, "SKU1001");

            //MessageBox.Show(teste.ToString());
            //getDemand();
            //getCurrentStock();
            //getItems();
            //getPlanningResources();
            //calculateNetRequirements();
            exportData();
            return 0; 
        }

        public int getDemand()
        {
                                  
            int DemandLength = sharedPreactor.RecordCount(tblNonAggDemand);
            try
            {
                for (int i = 1; i <= DemandLength; i++)
                {
                    Demand Demand = new Demand();
                    Demand.ItemCode = sharedPreactor.ReadFieldString(tblNonAggDemand, clnDemandItemCode, i);
                    Demand.OrderDate = sharedPreactor.ReadFieldDateTime(tblNonAggDemand, clnDemandOrderDate, i);
                    Demand.Quantity = sharedPreactor.ReadFieldDouble(tblNonAggDemand, clnDemandQuantity, i);
                    DemandList.Add(Demand);
                }
                //MessageBox.Show("Sucesso!");
            }
            catch
            {
                MessageBox.Show("Erro!");
            }              

            return 0;
        } 

        public int getCurrentStock()
        {
            

            int StockLength = sharedPreactor.RecordCount(tblStock);
            

            for (int i = 1; i <= StockLength; i++)
            {
                Stock Stock = new Stock();
                //Stock.Type = preactor.ReadFieldString(tblStock, "Type", i);
                Stock.ItemCode = sharedPreactor.ReadFieldString(tblStock, clnStockItemCode, i);
                Stock.ProdnDate = sharedPreactor.ReadFieldDateTime(tblStock, clnStockProdnDate, i);
                Stock.Qty = sharedPreactor.ReadFieldDouble(tblStock, clnStockQuantity, i);
                StockList.Add(Stock);
            }

            return 0;

        }

        public int getPlanningResources()
        {

            int ResourcesLength = sharedPreactor.RecordCount(tblPlanningResources);
            int DemandLength = sharedPreactor.RecordCount(tblNonAggDemand);
            var dates = DemandList.Select(x => x.OrderDate).Distinct();
            for (int i = 1; i <= ResourcesLength; i++)
            {
                foreach(var date in dates)
                {
                    
                    Resource Resource = new Resource();
                    Resource.ResourceName = sharedPreactor.ReadFieldString(tblPlanningResources, clnResourceName, i);
                    Resource.AvailableCapacityPeriod = 40;
                    Resource.DatePeriod = date;
                    ResourceList.Add(Resource);
                }
               
            }

            return 0;

        }


        public int getItems()
        {
            

            int ItemsLength = sharedPreactor.RecordCount(tblItem);
            

            for (int i = 1; i <= ItemsLength; i++)
            {
                sharedPreactor.ReadFieldString(tblItem, clnItemsItemLevel, i);
                if (sharedPreactor.ReadFieldString(tblItem, clnItemsItemLevel, i) == "Finished Product")
                {
                    Item Item = new Item();
                    Item.ItemCode = sharedPreactor.ReadFieldString(tblItem, clnItemsItemCode, i);
                    Item.ItemDesc = sharedPreactor.ReadFieldString(tblItem, clnItemsItemDesc, i);
                    Item.ItemLevel = sharedPreactor.ReadFieldString(tblItem, clnItemsItemLevel, i);
                    Item.CapacityUoM = sharedPreactor.ReadFieldString(tblItem, clnItemsCapacityUoM, i);
                    Item.PlanningResourceGroup = sharedPreactor.ReadFieldString(tblItem, clnItemsPlanningResourceGroup, i);
                    Item.ReorderMultiple = sharedPreactor.ReadFieldDouble(tblItem, clnItemsReorderMultiple, i);
                    Item.MinimumReorderMultiple = sharedPreactor.ReadFieldDouble(tblItem, clnItemsMinimumReorderMultiple, i);
                    Item.MinimumCoverDays = sharedPreactor.ReadFieldDouble(tblItem, clnItemsMinimumCoverDays, i);
                    Item.TargetCoverDays = sharedPreactor.ReadFieldDouble(tblItem, clnItemsTargetCoverDays, i);
                    Item.MaximumCoverDays = sharedPreactor.ReadFieldDouble(tblItem, clnItemsMaximumCoverDays, i);
                    ItemsList.Add(Item);
                }
            }

            return 0;
        }

        public int calculateNetRequirements()
        {

            getItems();
            getDemand();
            getCurrentStock();
            getPlanningResources();
            for (int i = 0; i < DemandList.Count; i++)
            {
                MPSResults MPSItem = new MPSResults();
                string currentItemCode = MPSItem.ItemCode = DemandList[i].ItemCode;
                double grossRequirements = MPSItem.GrossRequirements = DemandList[i].Quantity;
                double initialInventory = MPSItem.BeggingStock = getItemStock(currentItemCode, DemandList[i].OrderDate);
                double minimumLoSize = MPSItem.MinimumReorderMultiple = 123; // ajustar

                MPSItem.DemandDate = DemandList[i].OrderDate;
                MPSItem.NetRequirements = Math.Max((grossRequirements - initialInventory), 0); 

                MPSList.Add(MPSItem);
            }

            //var ResourcePeriodsQuery = from resource in ResourceList
            //                                   from demand in DemandList
            //                                   select new { resource.ResourceName, resource.AvailableCapacityPeriod, demand.OrderDate };
            //ResourceList.Clear();
            //foreach (var resource in ResourcePeriodsQuery){
            //    Resource ResourcePeriods = new Resource();
            //    ResourcePeriods.ResourceName = resource.ResourceName;
            //    ResourcePeriods.DatePeriod = resource.OrderDate;
            //    ResourcePeriods.AvailableCapacityPeriod = resource.AvailableCapacityPeriod;
                
            //    ResourceList.Add(ResourcePeriods);
            //}
            return 0;
        }

        public double getItemStock(string itemCode, DateTime stockDate)
        {
        
            double itemBeggingStock = 0;
            foreach (Stock stock in StockList)
            {
                string localStock = stock.ItemCode;
                DateTime localDate = stock.ProdnDate;
                if (localStock == itemCode && localDate == stockDate)
                {
                    itemBeggingStock = stock.Qty;
                    break;
                }
                else
                {
                    itemBeggingStock = 0;
                }
               
            }
            return itemBeggingStock;
        }

        public int readExportData()
        {
            for(int i = 1; i <= sharedPreactor.RecordCount(tblDemand); i++)
            {
                MPSExport MPSExportLine = new MPSExport();
                MPSExportLine.ItemCode = sharedPreactor.ReadFieldString(tblDemand, clnDemandCode, i);
                MPSExportLine.CapacityUsed = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandCapacityUsed, i);
                MPSExportLine.Period = sharedPreactor.ReadFieldString(tblDemand, clnDemandDate, i);
                MPSExportLine.InitialInventory = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandOpeningStock, i);
                MPSExportLine.GrossRequirements = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandDemand, i);
                MPSExportLine.MPS = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandMPS, i);
                MPSExportLine.PlanningResource = sharedPreactor.ReadFieldString(tblDemand, clnDemandPlanningResource, i);
                MPSExportLine.TargetStock = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandTargetStock, i);
                MPSExportLine.ClosingStock = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandClosingStock, i);
                MPSExportLine.MinStock = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandMinStock, i);
                MPSExportLine.MinDaysCover = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandMinDaysCover, i);
                MPSExportLine.TargetDaysCover = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandTargetDaysCover, i);
                MPSExportLine.TotalDaysCover = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandTotalDaysCover, i);
                MPSExportList.Add(MPSExportLine);
            }
            
            return 0;
        }

        public void exportData()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            readExportData();
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xls", ValidateNames = true })
            {
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                Workbook wb = app.Workbooks.Add(XlSheetType.xlWorksheet);
                Worksheet ws = (Worksheet)app.ActiveSheet;
                app.Visible = false;
                List<string> props = typeof(MPSExport).GetProperties().Select(f => f.Name).ToList();
                int i = 1;
                foreach(var prop in props)
                {
                    ws.Cells[1, i] = prop;
                    
                    i++;

                }
                int j = 2;
                foreach (MPSExport MPSExport in MPSExportList)
                {
                    ws.Cells[j, 1] = MPSExport.ItemCode.ToString();
                    ws.Cells[j, 2] = MPSExport.CapacityUsed.ToString();
                    ws.Cells[j, 3] = MPSExport.Period.ToString();
                    ws.Cells[j, 4] = MPSExport.OnHand.ToString();
                    ws.Cells[j, 5] = MPSExport.InitialInventory.ToString();
                    ws.Cells[j, 6] = MPSExport.GrossRequirements.ToString();
                    ws.Cells[j, 7] = MPSExport.StandardLotSize.ToString();
                    ws.Cells[j, 8] = MPSExport.NetRequirements.ToString();
                    ws.Cells[j, 9] = MPSExport.MPS.ToString();
                    ws.Cells[j, 10] = MPSExport.TargetStock.ToString();
                    ws.Cells[j, 11] = MPSExport.MinStock.ToString();
                    ws.Cells[j, 12] = MPSExport.ClosingStock.ToString();
                    ws.Cells[j, 13] = MPSExport.MinDaysCover.ToString();
                    ws.Cells[j, 14] = MPSExport.TargetDaysCover.ToString();
                    ws.Cells[j, 15] = MPSExport.TotalDaysCover.ToString();
                    ws.Cells[j, 16] = MPSExport.PlanningResource.ToString();
                    ws.Cells[j, 17] = MPSExport.RequirementsMet.ToString();
                    ws.Cells[j, 18] = MPSExport.RequirementsNotMet.ToString();
                    ws.Cells[j, 19] = MPSExport.ServiceLevel.ToString();
                    ws.Cells[j, 20] = MPSExport.AverageServiceLevelPeriod.ToString();
                    ws.Cells[j, 21] = MPSExport.EndingInventory.ToString();
                    ws.Cells[j, 22] = MPSExport.AverageInventoryPeriod.ToString();
                    ws.Cells[j, 23] = MPSExport.TotalAverageInventoryPeriod.ToString();
                    ws.Cells[j, 24] = MPSExport.BelowSafetyInvetory.ToString();
                    ws.Cells[j, 25] = MPSExport.BelowSafetyInvetoryPeriod.ToString();
                    j++;
                }
                wb.SaveAs("MPSFile", XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                app.Quit();
                MessageBox.Show("Seu arquivo foi exportado com sucesso!", "Mensagem");
            }
        }

        

        
        

  
    //public dynamic getItemData(string table = "", string field = "", string itemCode = "")
    //{
    //    int fmtTable = sharedPreactor.GetFormatNumber(table);
    //    int fmtField = sharedPreactor.GetFieldNumber(fmtTable, field);
    //    FormatFieldPair tablePair = new FormatFieldPair(fmtTable, fmtField);
    //    int itemFormatRecord = sharedPreactor.FindMatchingRecord(tablePair, 0, itemCode);
    //    PreactorFieldType fieldType = sharedPreactor.GetFieldType(tablePair);
    //    dynamic itemData = null;

    //    switch (fieldType)
    //    {
    //        case PreactorFieldType.String:
    //            itemData = sharedPreactor.ReadFieldString(tablePair, itemFormatRecord);
    //            break;
    //        case PreactorFieldType.Integer:
    //            itemData = sharedPreactor.ReadFieldInt(tablePair, itemFormatRecord); 
    //            break;
    //        case PreactorFieldType.DateTime:
    //            itemData = sharedPreactor.ReadFieldDateTime(tablePair, itemFormatRecord);
    //            break;
    //        case PreactorFieldType.Real:
    //            itemData = sharedPreactor.ReadFieldDouble(tablePair, itemFormatRecord);
    //            break;
    //        case PreactorFieldType.FreeFormatString:
    //            itemData = sharedPreactor.ReadFieldBool(tablePair, itemFormatRecord);
    //            break;
    //        default:

    //            break;
    //    }

    //    return itemData;
    //}


}
}
