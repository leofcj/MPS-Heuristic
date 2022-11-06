using System;
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
        IList<NonAggregateDemand> NonAggDemandList = new List<NonAggregateDemand>();
        IList<Stock> StockList = new List<Stock>();
        IList<Resource> ResourceList = new List<Resource>();
        IList<Demand> DemandList = new List<Demand>();
        IList<Item> ItemsResourceCount = new List<Item>();
        IList<Demand> MPSList = new List<Demand>();
        IList<MPSExport> MPSExportList = new List<MPSExport>();
        IList<ItemsResource> ItemsResourceData = new List<ItemsResource>();
        public int genMPS(ref PreactorObj preactorComObject, ref object pespComObject)
        {
            sharedPreactor = PreactorFactory.CreatePreactorObject(preactorComObject);

            //MessageBox.Show(teste.ToString());
            //getDemand();
            //getCurrentStock();
            //getItems();
            //getPlanningResources();
            getItems();
            getPlanningResources();
            calculateNetRequirements();
            createGridControl();
            getItemsResourceData();
            calculateMPS();
            //exportData();
            sharedPreactor.Planner.RefreshPlannerGrid();


            return 0; 
        }


        // calcula as necessidades líquidas para iniciar o processo de criação de MPS
        public int calculateNetRequirements()
        {

            // Net Req = Min(Mult(Min(Max([gross - (initial inv + subcont)],0), min lot size), standard lot size), max inv level)  
            //calcula os estoques para abrir o initial inventory
            sharedPreactor.Planner.CalculateStock();

            string lastCode = null;
            int demandLength = sharedPreactor.RecordCount(tblDemand);
            for (int i = 1; i <= demandLength; i++)
            {

                double initialInventory = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandOpeningStock, i);
                if (initialInventory < 0)
                {
                    sharedPreactor.WriteField(tblDemand, clnDemandOpeningStock, i, 0);
                    initialInventory = 0;
                }
                string currentCode = sharedPreactor.ReadFieldString(tblDemand, clnDemandCode, i);
                double currentNetRequirements = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandNetRequirements, i);
                if (initialInventory >= 0 && currentCode != lastCode && currentNetRequirements == 0)
                {
                    double subcontracted = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandSubcontracted, i);
                    double grossRequirements = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandDemand, i);
                    double minimumLotSize = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandMinimumLotSize, i);
                    double standardLotSize = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandReorderMultiple, i);
                    double maximumInventoryLevel = sharedPreactor.ReadFieldDouble(tblDemand, clndDemandMaximumInventoryLevel, i);
                    double safetyInventory = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandSafetyInventory, i);
                    double val1 = Math.Max(safetyInventory + grossRequirements - initialInventory - subcontracted, 0);
                    double val2 = Math.Max(val1, minimumLotSize);
                    double val3 = Math.Ceiling(val2 / standardLotSize) * standardLotSize;
                    double val4 = Math.Min(val3, maximumInventoryLevel);
                    double netRequirements = val4;

                    FormatFieldPair demandNetRequirements = new FormatFieldPair(sharedPreactor.GetFormatNumber(tblDemand), sharedPreactor.GetFieldNumber(tblDemand, clnDemandNetRequirements));
                    sharedPreactor.WriteField(demandNetRequirements, i, netRequirements);
                }

                lastCode = currentCode;


            }
            sharedPreactor.Planner.RefreshPlannerGrid();
            return 0;
        }



        public int getItemsResourceCount(string itemCode)
        {
            int itemLength = sharedPreactor.RecordCount(tblItem);
            ItemsResourceCount.Clear();
            int count = 0;
            for (int i = 1; i <= itemLength; i ++)
            {
                MatrixDimensions size = sharedPreactor.MatrixFieldSize(tblItem, clnItemsPlanningResourceData, i);
                Item Item = new Item();
                Item.ItemCode = sharedPreactor.ReadFieldString(tblItem, clnItemsItemCode, i);
                if (itemCode == Item.ItemCode.ToString())
                {
                    count = size.X;
                }
                Item.ResourceCount = size.X;
                ItemsResourceCount.Add(Item);
                
            }
            ItemsResourceCount = ItemsResourceCount.OrderBy(x => x.ResourceCount).ToList();

            return count;
        }

        



        public int getNonAggDemand()
        {

            int DemandLength = sharedPreactor.RecordCount(tblNonAggDemand);
            try
            {
                for (int i = 1; i <= DemandLength; i++)
                {
                    NonAggregateDemand Demand = new NonAggregateDemand();
                    Demand.ItemCode = sharedPreactor.ReadFieldString(tblNonAggDemand, clnDemandItemCode, i);
                    Demand.OrderDate = sharedPreactor.ReadFieldDateTime(tblNonAggDemand, clnDemandOrderDate, i);
                    Demand.Quantity = sharedPreactor.ReadFieldDouble(tblNonAggDemand, clnDemandQuantity, i);
                    NonAggDemandList.Add(Demand);
                }
                //MessageBox.Show("Sucesso!");
            }
            catch
            {
                MessageBox.Show("Erro!");
            }

            return 0;
        }



        public int getPlanningResources()
        {
            getNonAggDemand();
            
            int ResourcesLength = sharedPreactor.RecordCount(tblPlanningResources);

            var dates = NonAggDemandList.Select(x => x.OrderDate).Distinct();
            for (int i = 1; i <= ResourcesLength; i++)
            {
                foreach (var date in dates)
                {

                    Resource Resource = new Resource();
                    Resource.ResourceName = sharedPreactor.ReadFieldString(tblPlanningResources, clnResourceName, i);
                    Resource.AvailableCapacityPeriodInWeek = 40;
                    Resource.AvailableCapacityPeriodInMonth = 170;
                    Resource.OriginalCapacityPeriodInWeek = 40;
                    Resource.OriginalCapacityPeriodInMonth = 170;
                    Resource.OvertimePercent = 20;
                    Resource.DatePeriod = date;
                    ResourceList.Add(Resource);
                }

            }

            return 0;

        }


        public double getResourceAvaliableCapacity(string resource, DateTime date)
        {
            int resourcesLength = ResourceList.Count();
            double availableCapacity = 0;
            for (int i = 0; i < resourcesLength; i++)
            {
                if (ResourceList[i].ResourceName == resource && ResourceList[i].DatePeriod == date)
                {
                    availableCapacity = ResourceList[i].AvailableCapacityPeriodInWeek;
                    break;
                }
            }
            return availableCapacity;
        }

        public void setResourceStats(string resource, DateTime date, double newCapacity)
        {
            int resourcesLength = ResourceList.Count();

            for (int i = 0; i < resourcesLength; i++)
            {
                if (ResourceList[i].ResourceName == resource && ResourceList[i].DatePeriod == date)
                {
                    ResourceList[i].AvailableCapacityPeriodInWeek = newCapacity;
                    break;
                }
            }
            
        }

        public void setDemandFit(string itemCode, DateTime date, int fitsInFull, int fitsInPartial = 0)
        {
            int demandLength = DemandList.Count();
            for(int i = 0; i < demandLength; i++)
            {
                if(DemandList[i].ItemCode == itemCode && DemandList[i].DemandDate == date)
                {
                    DemandList[i].ItemFitsInFull = fitsInFull;
                    DemandList[i].ItemFitsInPartial = fitsInPartial;
                }
            }
        }


        public void getItemsResourceData()
        {
            int itemsLength = sharedPreactor.RecordCount(tblItem);

            for (int i = 1; i <= itemsLength; i++)
            {
                
                MatrixDimensions size = sharedPreactor.MatrixFieldSize(tblItem, clnItemsPlanningResourceData, i);
                for (int j = 1; j <= size.X; j++)
                {
                    ItemsResource ItemResources = new ItemsResource();
                    ItemResources.ItemCode = sharedPreactor.ReadFieldString(tblItem, clnItemsItemCode, i);
                    ItemResources.ResourceCode = sharedPreactor.ReadFieldString(tblItem, clnItemsPlanningResourceData, i, j);
                    ItemResources.RateperHour = sharedPreactor.ReadFieldDouble(tblItem, clnItemsResSpecificRateperHour, i, j);
                    ItemsResourceData.Add(ItemResources);
                }                             
            }
        }
        public double getResourceRate(string itemCode, string resource)
        {
            int resourceDataLength = ItemsResourceData.Count();
            double rate = 0;
            for (int i = 0; i < resourceDataLength; i++)
            {   
                if(ItemsResourceData[i].ItemCode == itemCode && ItemsResourceData[i].ResourceCode == resource)
                {
                    rate = ItemsResourceData[i].RateperHour;
                    break;
                }

            }
            return rate;
        }


        public int createGridControl()
        {
            int demandLength = sharedPreactor.RecordCount(tblDemand);
            for (int i = 1; i <= demandLength; i++)
            {
                Demand Demand = new Demand();
                Demand.NetRequirements = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandNetRequirements, i);
                if(Demand.NetRequirements > 0)
                {
                    Demand.Number = sharedPreactor.ReadFieldInt(tblDemand, clnDemandNumber, i);
                    Demand.ItemCode = sharedPreactor.ReadFieldString(tblDemand, clnDemandCode, i);
                    Demand.DemandDate = sharedPreactor.ReadFieldDateTime(tblDemand, clnDemandDate, i);
                    Demand.Resource = sharedPreactor.ReadFieldString(tblDemand, clnDemandPlanningResource, i);
                    Demand.CapacityUsed = sharedPreactor.ReadFieldDouble(tblDemand, clnDemandCapacityUsed, i);
                    Demand.ResourceCount = getItemsResourceCount(Demand.ItemCode.ToString());
                    Demand.ItemFitsInFull = 0;
                    Demand.ItemFitsInPartial = 0;                     
                    DemandList.Add(Demand);
                }
                DemandList = DemandList.OrderBy(x => x.DemandDate).ThenBy(c => c.ResourceCount).ThenByDescending(n => n.NetRequirements).ToList();
            }
                return 0;
        }

        public string chooseResource(string itemCode, DateTime date)
        {
            var resources = from resourceData in ItemsResourceData
                           where resourceData.ItemCode == itemCode
                           select resourceData;
            string choosenResource = "";
            double auxMPS = 0;
            foreach(var resource in resources)
            {
                string currentResource = resource.ResourceCode;
                double currentRate = resource.RateperHour;
                double currentAvailableCapacity = getResourceAvaliableCapacity(currentResource, date);
                double possibleMPS = currentRate * currentAvailableCapacity;
                if (auxMPS <= possibleMPS)
                {
                    auxMPS = possibleMPS;
                    choosenResource = currentResource;
                }
            }

            return choosenResource;
        }


        public int calculateMPS()
        {
            int demandListLength = DemandList.Count();

            for (int i = 0; i < demandListLength; i ++)
            {
                int fitsInFull = DemandList[i].ItemFitsInFull;

                if(fitsInFull != 1)
                {
                    //local variables
                    string currentResource = DemandList[i].Resource;
                    string currentItemCode = DemandList[i].ItemCode;
                    double currentNetRequirements = DemandList[i].NetRequirements;
                    DateTime currentDate = DemandList[i].DemandDate;
                    string choosenResource = chooseResource(currentItemCode, currentDate);
                    double resourceAvailableCapacity = getResourceAvaliableCapacity(currentResource, currentDate);
                    double resourceRate = getResourceRate(currentItemCode, currentResource);
                    double resourceNecessaryCapacity = currentNetRequirements / resourceRate;
                    double resourceOvertime = 20;
                    double newCapacity = resourceAvailableCapacity - resourceNecessaryCapacity;
                    int record = sharedPreactor.FindMatchingRecord(tblDemand, clnDemandNumber, 0, DemandList[i].Number);
                    double minimumLotSize = ItemsList.ToList().Find(x => (x.ItemCode == currentItemCode)).MinimumReorderMultiple;
                    double possibleMPS = resourceAvailableCapacity * resourceRate * (100 + resourceOvertime) / 100;

                    if (choosenResource == currentResource)
                    {
                        if(resourceNecessaryCapacity <= resourceAvailableCapacity)
                        {
                            DemandList[i].MPS = DemandList[i].NetRequirements;
                            sharedPreactor.WriteField(tblDemand, clnDemandMPS, record, DemandList[i].MPS);
                            
                            setResourceStats(choosenResource, currentDate, newCapacity);
                            if (newCapacity == 0)
                            {
                                setDemandFit(currentItemCode, currentDate, 1, 0);
                            }
                        }
                        else if (resourceNecessaryCapacity > resourceAvailableCapacity)
                        {
                            var originalCapacity = ResourceList.ToList().Find(r => (r.ResourceName == currentResource) && (r.DatePeriod.Equals(currentDate))).OriginalCapacityPeriodInWeek;

                            if (resourceAvailableCapacity == originalCapacity)
                            {
                                DemandList[i].MPS = possibleMPS;
                                sharedPreactor.WriteField(tblDemand, clnDemandMPS, record, DemandList[i].MPS);
                                setResourceStats(choosenResource, currentDate, 0);
                                setDemandFit(currentItemCode, currentDate, 1, 0);

                            }
                            
                            else if (minimumLotSize <= possibleMPS)
                            {
                                DemandList[i].MPS = possibleMPS;
                                sharedPreactor.WriteField(tblDemand, clnDemandMPS, record, DemandList[i].MPS);
                                setResourceStats(choosenResource, currentDate, 0);
                                setDemandFit(currentItemCode, currentDate, 0, 1);
                            }

                        }
                       
                    }
                }

                

                //if (resourceNecessaryCapacity <= resourceAvailableCapacity && fitsInFull != 1)
                //{
                //    DemandList[i].MPS = DemandList[i].NetRequirements;
                //    sharedPreactor.WriteField(tblDemand, clnDemandMPS, record, DemandList[i].MPS);
                //    double newCapacity = resourceAvailableCapacity - resourceNecessaryCapacity;
                //    setResourceStats(choosenResource, currentDate, newCapacity);
                //    if (newCapacity >= 0)
                //    {
                //        setDemandFit(currentItemCode, currentDate, 1, 0);
                //    }
                //}

                //else if(resourceNecessaryCapacity > resourceAvailableCapacity && fitsInFull != 1)
                //{
                //    double possibleMPS = resourceAvailableCapacity * resourceOvertime/100 * resourceRate;
                //    DemandList[i].MPS = possibleMPS;
                //    sharedPreactor.WriteField(tblDemand, clnDemandMPS, record, DemandList[i].MPS);

                //}



            }
            sharedPreactor.Planner.CalculateStock();
            return 0;
        }


        //public int getCurrentStock()
        //{


        //    int StockLength = sharedPreactor.RecordCount(tblStock);


        //    for (int i = 1; i <= StockLength; i++)
        //    {
        //        Stock Stock = new Stock();
        //        //Stock.Type = preactor.ReadFieldString(tblStock, "Type", i);
        //        Stock.ItemCode = sharedPreactor.ReadFieldString(tblStock, clnStockItemCode, i);
        //        Stock.ProdnDate = sharedPreactor.ReadFieldDateTime(tblStock, clnStockProdnDate, i);
        //        Stock.Qty = sharedPreactor.ReadFieldDouble(tblStock, clnStockQuantity, i);
        //        StockList.Add(Stock);
        //    }

        //    return 0;

        //}




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

        //public int calculateNetRequirements()
        //{

        //    //Net Req = Min(Mult(Min(Max([gross - (initial inv + subcont)],0), min lot size), standard lot size), max inv level)    

        //    getItems();
        //    getDemand();
        //    getCurrentStock();
        //    getPlanningResources();
        //    for (int i = 0; i < NonAggDemandList.Count; i++)
        //    {
        //        MPSResults MPSItem = new MPSResults();
        //        string currentItemCode = MPSItem.ItemCode = NonAggDemandList[i].ItemCode;
        //        double grossRequirements = MPSItem.GrossRequirements = NonAggDemandList[i].Quantity; 
        //        double initialInventory = MPSItem.BeggingStock = getItemStock(currentItemCode, NonAggDemandList[i].OrderDate);
        //        double minimumLoSize = MPSItem.MinimumReorderMultiple = 123; // ajustar
        //        MPSItem.DemandDate = NonAggDemandList[i].OrderDate;
        //        MPSItem.NetRequirements = Math.Max((grossRequirements - initialInventory), 0); 
        //        MPSList.Add(MPSItem);
        //    }



        //    return 0;
        //}

        //public double getItemStock(string itemCode, DateTime stockDate)
        //{

        //    double itemBeggingStock = 0;
        //    foreach (Stock stock in StockList)
        //    {
        //        string localStock = stock.ItemCode;
        //        DateTime localDate = stock.ProdnDate;
        //        if (localStock == itemCode && localDate == stockDate)
        //        {
        //            itemBeggingStock = stock.Qty;
        //            break;
        //        }
        //        else
        //        {
        //            itemBeggingStock = 0;
        //        }

        //    }
        //    return itemBeggingStock;
        //}





        //Funcao de exportar dados abaixo. 

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
