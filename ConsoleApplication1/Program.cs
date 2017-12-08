using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Marshal = System.Runtime.InteropServices;
using System.Security.Principal;
using NLog;
using System.Xml;
using System.Data;

namespace FlowPlanConstruction
{
    class Program
    {
        private static Logger log = LogManager.GetCurrentClassLogger();

        //section toggles
        private const bool debugMode = true;
        public const bool runArchive = true;
        public const bool runLaborPlan = true;
        public const bool runCustom = true;

        //application flags
        public const bool visibilityFlag = false;
        public const bool alertsFlag = false;

        //path information
        public const string chargeDataSrcPath = @"\\cfc1afs01\Operations-Analytics\RAW Data\chargepattern.csv";
        public const string masterFlowPlanLocation = @"\\CFC1AFS01\Operations-Analytics\Projects\Flow Plan\Outbound Flow Plan v2.0(MCRC).xlsm";
        public const string laborPlanLocationation = @"\\chewy.local\bi\BI Community Content\Finance\Labor Models\";
        public const string whXmlList = @"\\cfc1afs01\Operations-Analytics\Projects\Flow Plan\Release\Warehouses.xml";
        public const string masterPreShiftCheckList = @"\\cfc1afs01\Operations-Analytics\Projects\Documentation\Standard Work\";

        //shiift defenition
        public static string[] shifts = { "Days", "Nights", "Blank" };

        //rate defaults
        public static double[] avp1TPHHourlyGoalDays = { 0.80, 1.12, 1.12, 0.80, 1.08, 0.96, 1.12, 0.80, 1.12, 1.04 };
        public static double[] avp1TPHHourlyGoalNights= { 0.80, 1.12, 1.12, 0.80,  1.08, 0.96, 1.12, 0.80, 1.12, 1.04 };
        public static double[] defaultTPHHourlyGoalDays = { .90, 1, .90,1,1,.7,.9, 1,.9, 1,.7};
        public static double singlesPercentSplit = .33;
        public static double shiftLength = 9.5;
        public static double breakModifier = .7;
        public static double UPC = 2.5;
        public static double defaultMultiRate = 47;

        //add a way to edit the XML file external to this program could be stand alone configuration that would allow user to kick off a run of the constructor with new perimeters....?
        //sql query

        public static string prodConnectionString = "Data source=WMSSQL-READONLY.chewy.local;"
               + "Initial Catalog=AAD;"
               + "Persist Security Info=True;"
               + "Trusted_connection=true";
        public static string chargeQuery = "Use AAD go with item_order_quantity as (SELECT o.order_number,o.wh_id,COUNT(DISTINCT d.item_number) item_count,sum(d.qty) unit_qty FROM AAD.dbo.t_order o left join AAD.dbo.t_order_detail d on o.order_number = d.order_number and o.wh_id = d.wh_id where o.arrive_date > cast(getdate() - 60 as date) group by o.order_number,o.wh_id) select o.wh_id, cast(o.arrive_date as date) arrive_date, datepart(hour,o.arrive_date) arrive_hour, case when q.item_count = 1 then 'SINGLES' else 'MULTIS' end order_type, o.che_route, count(distinct o.order_number) order_ct, sum(q.unit_qty) unit_qty from AAD.dbo.t_order o left join item_order_quantity q on q.order_number = o.order_number and q.wh_id = o.wh_id where o.type_id = '31' and o.arrive_date > Cast(getdate() - 60 as date) group by cast(o.arrive_date as date), o.wh_id, datepart(hour,o.arrive_date), o.che_route, case when q.item_count = 1 then 'SINGLES' else 'MULTIS' end order by  cast(o.arrive_date as date), o.wh_id, datepart(hour,o.arrive_date), o.che_route, case when q.item_count = 1 then 'SINGLES' else 'MULTIS' end";

        static void Main(string[] args)
        {
            //setup log file information
            string user = WindowsIdentity.GetCurrent().Name;

            log.Info("Application started by {0}", user);

            bool runLaborPlanPopulate = false;

            double[] laborplaninfo = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

            List<Warehouse> warehouseList = new List<Warehouse>();

            warehouseList = getWHList();
            try
            {
                foreach (Warehouse wh in warehouseList)
                {
                    if (!debugMode )//&& wh.Name =="CFC1")
                    {
                        log.Info("Debug mode disabled checking process flags.");
                        if (runLaborPlan)
                        {
                            log.Info("Labor Plan process Initalized for {0}", wh.Name);
                            wh.laborPlanInforRows = obtainLaborPlanInfo(wh.laborPlanInforRows, wh.Name, laborplaninfo, ref runLaborPlanPopulate);
                        }

                        if (runArchive)
                        {
                            log.Info("Archive process Initalized for {0}", wh.Name);
                            CleanUpDirectories(wh.blankCopyLoc, wh.archiveLoc, wh.daysRate, wh.nightsRate);
                        }

                        if (runCustom)
                        {
                            log.Info("Flow Plan Customization process Initalized for {0}", wh.Name);
                            foreach (string shift in shifts)
                            {
                                customizeFlowPlan(wh.blankCopyLoc, wh.Name, shift, runLaborPlanPopulate, laborplaninfo, wh.DistroList, wh.HandoffPercent, wh.VCPUWageRate, wh.DeadMan, wh.daysRate, wh.nightsRate, wh.tphDistro, wh.laborInfoPop);
                                if (wh.PreShiftFlag)
                                    CustomizePreShift(laborplaninfo, shift, wh.blankCopyLoc, wh.Name, wh.preshiftInfoPop);
                            }
                        }
                    }
                }
        }
            catch (Exception e)
            {
                CrashshutDownApplication(e);
             }


            GracefulExit(warehouseList);
            
        }
        private static void CleanUpDirectories(string locToArchive, string archiveLocation, double[] DayRates, double[] NightsRates)
        {
            string Lvl1rollup = @"\\CFC1AFS01\Operations-Analytics\Projects\Flow Plan\RollUpInfo\LVL1rollup.xlsx";

            //Define Excel Application
            Excel.Application archiveApp = null;
            archiveApp = new Excel.Application();
            archiveApp.Visible = visibilityFlag;
            archiveApp.DisplayAlerts = alertsFlag;

            //set up workbooks
            Excel.Workbooks archiveworkbooks = null;
            archiveworkbooks = archiveApp.Workbooks;
            Excel.Workbook collectionSummary = null;
            Excel.Workbook currentProcessableFile = null;


            try
            {
                collectionSummary = archiveworkbooks.Open(Lvl1rollup, false, false);
            }catch
            {
                log.Warn("Unable to open " + Lvl1rollup);

            }

            DateTime modifiedDate = new DateTime(1900, 01, 01);
            string lastmodified = "";
            string filename2 = "";
            string archivePathFile = "";
            string archivePathloca = "";

            string[] blankcopies = Directory.GetFiles(locToArchive);
            foreach (string copy in blankcopies)
            {
                modifiedDate = System.IO.File.GetLastWriteTime(copy);
                lastmodified = modifiedDate.ToString("yyyy-MM-dd");
                lastmodified = lastmodified.Replace('-', '\\');
                filename2 = System.IO.Path.GetFileName(copy);
                archivePathloca = System.IO.Path.Combine(archiveLocation, lastmodified);
                var dir = new DirectoryInfo(archivePathloca);
                archivePathFile = System.IO.Path.Combine(archivePathloca, filename2);

                try
                {
                    currentProcessableFile = archiveworkbooks.Open(copy, false, true);
                }
                catch
                {
                    log.Warn("Could not open file at: {0} ", copy);
                    continue;
                }

                var emptycheck = currentProcessableFile.Worksheets.Item["SOS"].cells[9, 5].value;

                if(emptycheck == null)
                {
                    log.Info("File empty removing");
                    currentProcessableFile.Close(false);
                    try
                    {
                        File.Delete(copy);
                        log.Info("Empyt File removed");
                    }catch
                    {
                        log.Warn("Unable to remove blank file: {0}", copy);
                    }
                }
                else
                {
                    if (dir.Exists)
                    {
                        log.Info("Directory confirmend at: " + dir.ToString());
                    }
                    else
                    {
                        log.Info("Directory created at: " + dir.ToString());
                        System.IO.Directory.CreateDirectory(dir.ToString());
                    }

                    try
                    {
                        gatherCustomInfo(currentProcessableFile, DayRates, NightsRates);
                        gatherRollUpInfo(collectionSummary, currentProcessableFile);
                        currentProcessableFile.Close(false);
                        log.Info("File Coppied from: " + copy + " to Archive");
                        System.IO.File.Copy(copy, archivePathFile, true);
                        System.IO.File.Delete(copy);
                    }
                    catch (Exception)
                    {
                        log.Warn(System.DateTime.Now + ":\t" + "File currently in use: " + copy);
                    }
                }
                
            }

            //clean up excel
            collectionSummary.SaveAs(Lvl1rollup);
            collectionSummary.Close();
            collectionSummary = null;
            archiveworkbooks = null;
            archiveApp.Quit();
            archiveApp = null;
        }

        public static void gatherCustomInfo(Excel.Workbook CurrentFile, double[] daysRate, double[] nightsRate)
        {
            string shift = CurrentFile.Worksheets.Item["SOS"].cells[5, 9].value;
            double[] rates = {0,0,0,0,0,0,0,0,0 };
            Excel.Worksheet HourlyST = CurrentFile.Worksheets.Item["Hourly ST"];

            for (int i = 0; i <= 8; i++)
            {
                rates[i] = HourlyST.Cells[i + 52, 4].value;
            }

            if(shift == "Days")
            {
                rates.CopyTo(daysRate, 0);
            }else if(shift == "Nights")
            {
                rates.CopyTo(nightsRate, 0);
            }
            HourlyST = null;


        }
        public static void gatherRollUpInfo(Excel.Workbook roll1, Excel.Workbook CurrentFile)
        {

            string[] lvl1rollupCellsSOS =
                {
                "I2", "I3", "I4", "I5", "C9", "C10", "C11", "C12",
                "C14", "C15", "C16", "C17", "E9", "E10", "E11",
                "E12", "G9", "G10","G11", "G12", "I9", "I10", "I11",
                "I12"
                }; // here we list ever cell we need informatino from for the version 2 or greater
            string[] lvl1rollupCellsEOS =
            {
                "D30", "D31", "D33", "D34", "D35","D39"
            };
            string[] lvl1rollupCellsEOSolderversion =
            {
                "E23","E24","E26","E27","E28", "E20"
            };
            string[] lvl1processableCellsSOS = { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
            string[] lvl1processableCellsEOS = { "", "", "", "", "", "" };

            Excel.Worksheet rollUpSheet = null;
            Excel.Worksheet SOS = null;
            Excel.Worksheet EOS = null;

            rollUpSheet = roll1.Sheets.Item[1];
            SOS = CurrentFile.Sheets.Item["SOS"];
            EOS = CurrentFile.Sheets.Item["EOS"];

            int RollUpLastRow = roll1.Worksheets[1].usedRange.Rows.count;
            string version = CurrentFile.Worksheets.Item["SOS"].cells(2, 9).value;

            if (version.Contains("2.0.1"))
            {
                //process newest version file
                lvl1rollupCellsSOS.CopyTo(lvl1processableCellsSOS, 0);
                lvl1rollupCellsEOS.CopyTo(lvl1processableCellsEOS, 0);
            }
            else if (version.Contains("1.9") || version.Contains("2.0"))
            {
                //process older version file
                lvl1rollupCellsSOS.CopyTo(lvl1processableCellsSOS, 0);
                lvl1processableCellsSOS[22] = "I8";
                lvl1rollupCellsEOSolderversion.CopyTo(lvl1processableCellsEOS, 0);
                SOS.Cells[17, 3].value = SOS.Cells[12, 3].value - SOS.Cells[22, 9].value;
            }
            else
            {
                Console.WriteLine("Incorrect Version Information: " + version);
                CurrentFile.Close(false);
            }

            //process all of the lvl1 roll up versions

            int i = 1;

            foreach (string cell in lvl1processableCellsSOS)
            {
                Excel.Range rng = SOS.Range[cell];
                roll1.Worksheets.Item[1].cells(RollUpLastRow + 1, i).value = rng.Value;
                i++;
            }
            foreach (string cell in lvl1processableCellsEOS)
            {
                Excel.Range rng = EOS.Range[cell];
                roll1.Worksheets.Item[1].cells(RollUpLastRow + 1, i).value = rng.Value;
                i++;
            }

            EOS = null;
            SOS = null;
            rollUpSheet = null;

        }
        private static int[] obtainLaborPlanInfo( int[]dprows, string wh, double[] laborplaninfo, ref bool laborPlanPopulate)
        {
            log.Info("Begin Gathereing infromation form " + wh + " Labor Plan");

            //Define Excel Application
            Excel.Application LaborPlanApp = null;
            LaborPlanApp = new Excel.Application();
            LaborPlanApp.Visible = visibilityFlag;
            LaborPlanApp.DisplayAlerts = alertsFlag;


            Excel.Workbooks Workbooks = null;
            Workbooks = LaborPlanApp.Workbooks;

            string[] laborplans = Directory.GetFiles(laborPlanLocationation);
            string laborplanfileName = "";
            Excel.Workbook laborPlanModel = null;
            int col = 0;
            col = System.DateTime.Today.DayOfYear;
            int i = 0;
            
            string[] rowlabels = { "Planned Total Show Hours (Days)", "Planned Throughput (Days)", "Planned Units Shipped (Days)", "Planned Supply Chain Units Ordered (Days)", "Planned Total Show Hours (Nights)", "Planned Throughput (Nights)", "Planned Units Shipped (Nights)", "Planned Supply Chain Units Ordered (Nights)", "Planned Total Show Hours (Days)", "Planned Throughput (Days)", "Planned Units Shipped (Days)", "Planned Supply Chain Units Ordered (Days)" };


            foreach (string laborplan in laborplans)
            {
                int r = 0;
                laborplanfileName = System.IO.Path.GetFileName(laborplan);
                if (laborplanfileName.Contains(wh) && laborplanfileName.Contains(".xls"))
                {

                    //open 21 DP
                    try
                    {
                        laborPlanModel = Workbooks.Open(laborplan, false, true);
                    }
                    catch
                    {
                        log.Warn("Unable to open labor plan at: " + laborplan);
                        laborPlanPopulate = false;
                    }
                    
                    Excel.Worksheet OBDP = laborPlanModel.Worksheets.Item["OB Daily Plan"];

                    //determine which rows to go look for
                    foreach (string label in rowlabels)
                    {
                        
                        if (r > 11)
                        {
                            log.Warn("Current Row Check Greater than number of verifications needed.. Shutting down Labor Plan Search");
                            laborPlanPopulate = false;
                        }
                        else
                        {
                            string rcheck = OBDP.Cells[dprows[r], 2].value; //Set the row check = to what we think the row should be
                            if (rcheck == label)
                            {
                                log.Info(wh + " Row location verified for " + label + " at row " + dprows[r]);

                            }
                            else
                            {
                                int OBDPusedRange = OBDP.UsedRange.Rows.Count;
                                string newrcheck = "";
                                for (int x = 1; x < OBDPusedRange; x++)
                                {
                                    newrcheck = OBDP.Cells[x, 2].value;
                                    if (newrcheck == label)
                                    {
                                        dprows[r] = x;
                                        log.Warn(wh + " Row location for " + label + " now found at row " + dprows[r]);
                                    }

                                }
                            }
                        }
                        r++;
                    }


                    foreach (int row in dprows)
                    {
                        if(i>11)
                        {
                            log.Warn("Current Row Check Greater than number of verifications needed.. Shutting down Labor Plan Search");
                            laborPlanPopulate = false;
                        }
                        //add information into array
                        if (i < 9)//collect current day
                        {
                            laborplaninfo[i] = OBDP.Cells[row, col + 2].value;//cloumn == 2 + number of days past this year

                        }
                        else//last 4 get from tomorrow for next day planning
                        {
                            laborplaninfo[i] = OBDP.Cells[row, col + 3].value;//cloumn == 2 + number of days past this year + 1 for tomorrow
                        }

                        i++;
                    }
                    OBDP = null;
                    laborPlanModel.Close();
                }
            }

            LaborPlanApp.Quit();
            Workbooks = null;
            LaborPlanApp = null;

            laborPlanPopulate = true;
            return dprows; // we want to add the data to the flow planners
        }

        private static void customizeFlowPlan(string newFileDirectory, string warehouse, string shift,bool laborPlanInfoAvaliable, double[] laborPlanInformation, string distrobutionList, double backlogHandoffPercent, double VCPUWageRate, int DeadMan, double[] daysRates, double[] nightsRate, double[] tphGoals, bool laborPlanPop)// this may need restructuring to make it more readable keep and eye on for future improvements
        {
            log.Info("Begining {0} {1} file customization process", warehouse, shift);
            //define Excel Application for the Creation Process
            Excel.Application CustomiseFlowPlanApplication = null;
            CustomiseFlowPlanApplication = new Excel.Application();
            CustomiseFlowPlanApplication.Visible = visibilityFlag;
            CustomiseFlowPlanApplication.DisplayAlerts = alertsFlag;
            
            //create the empty containers for the excel application
            Excel.Workbooks customizeFlowPlanWorkBookCollection = null;
            Excel.Workbook chargeDataSourceWB = null;
            Excel.Worksheet chargeDataSourceWKST = null;
            Excel.Workbook customFlowPlanDestinationWB = null;
            Excel.Worksheet customFlowPlanDestinationCHRGDATAWKST = null;
            Excel.Worksheet customFPDestinationSOSWKST = null;
            Excel.Worksheet customFlowPlanDestinationHOURLYTPHWKST = null;
            Excel.Worksheet customFlowPlanDestinationHourlySTWKST = null;
            Excel.Worksheet customFlowPlanDestinationMasterDataWKST = null;

            log.Info("Excel Empty containers created");

            //give the empty excel containers some meaning ----- need to add try catch blocks on src and dest file open----
            customizeFlowPlanWorkBookCollection = CustomiseFlowPlanApplication.Workbooks;
            chargeDataSourceWB = customizeFlowPlanWorkBookCollection.Open(chargeDataSrcPath, false, true);
            chargeDataSourceWKST = chargeDataSourceWB.Worksheets.Item[1];
            customFlowPlanDestinationWB = customizeFlowPlanWorkBookCollection.Open(masterFlowPlanLocation, false, false);
            customFlowPlanDestinationCHRGDATAWKST = customFlowPlanDestinationWB.Worksheets.Item["Charge Data"];
            customFPDestinationSOSWKST = customFlowPlanDestinationWB.Worksheets.Item["SOS"];
            customFlowPlanDestinationHOURLYTPHWKST = customFlowPlanDestinationWB.Worksheets.Item["Hourly TPH"];
            customFlowPlanDestinationHourlySTWKST = customFlowPlanDestinationWB.Worksheets.Item["Hourly ST"];
            customFlowPlanDestinationMasterDataWKST = customFlowPlanDestinationWB.Worksheets.Item["Master Data"];

            log.Info("Charge Data Source: {0}", chargeDataSrcPath);
            log.Info("Master Flow Plan Source: {0}", masterFlowPlanLocation);

            //set application calculation method
            CustomiseFlowPlanApplication.Calculation = Excel.XlCalculation.xlCalculationManual;

            double[] rates = {0,0,0,0,0,0,0,0,0 };
            //gather the sizes of ranges for each worksheet
            int sourceUsedRange = chargeDataSourceWKST.UsedRange.Rows.Count;
            int destinationUsedRange = customFlowPlanDestinationCHRGDATAWKST.UsedRange.Rows.Count;
            Excel.Range customFlowPlanStartPoint = customFlowPlanDestinationCHRGDATAWKST.Cells[2, 1]; //define the first cell below the column headers
            Excel.Range customFlowPlanEndPoint = customFlowPlanDestinationCHRGDATAWKST.Cells[destinationUsedRange, 6];
            Excel.Range customFlowPlanDestRange = customFlowPlanDestinationCHRGDATAWKST.Range[customFlowPlanStartPoint, customFlowPlanEndPoint];

            //empty the destination rage to avoid ghosting data add new data
            customFlowPlanDestRange.Value = "";

            log.Info("Charge Data cleared in master");

            customFlowPlanDestRange.Value = chargeDataSourceWKST.UsedRange.Value;

            //object test = chargeDataSourceWKST.UsedRange.Value;

            

            log.Info("New charge Data inserted into the master copy");

            chargeDataSourceWB.Close();

            //setup SOS Information
            customFPDestinationSOSWKST.Cells[3, 9].value = warehouse;
            customFPDestinationSOSWKST.Cells[4, 9].value = System.DateTime.Now.ToString("yyyy-MM-dd");
            customFPDestinationSOSWKST.Cells[5, 9].value = shift;

            log.Info("Start of shift information updated: Warehouse: {0} Date: {1} Shift: {2}", warehouse, System.DateTime.Now.ToString("yyyy-MM-dd"), shift);



            if(laborPlanInfoAvaliable && laborPlanPop)
            {
                switch (shift)
                {
                    case "Days":
                        customFPDestinationSOSWKST.Cells[10, 3].value = laborPlanInformation[0];//days hours
                        customFPDestinationSOSWKST.Cells[11, 3].value = laborPlanInformation[1];//days tph
                        customFPDestinationSOSWKST.Cells[9, 3].value = laborPlanInformation[2];//days shipped
                        customFPDestinationSOSWKST.Cells[12, 3].value = laborPlanInformation[3];//days ordered
                        customFPDestinationSOSWKST.Cells[15, 3].value = laborPlanInformation[4];//nights hours
                        customFPDestinationSOSWKST.Cells[16, 3].value = laborPlanInformation[5];//nights tph
                        customFPDestinationSOSWKST.Cells[14, 3].value = laborPlanInformation[6];//nights shipped
                        customFPDestinationSOSWKST.Cells[17, 3].value = laborPlanInformation[7];//nights ordered
                        break;
                    case "Nights":
                        customFPDestinationSOSWKST.Cells[15, 3].value = laborPlanInformation[8];//next days hours
                        customFPDestinationSOSWKST.Cells[16, 3].value = laborPlanInformation[9];//next days tph
                        customFPDestinationSOSWKST.Cells[14, 3].value = laborPlanInformation[10];//next days shipped
                        customFPDestinationSOSWKST.Cells[17, 3].value = laborPlanInformation[11];//next days ordered
                        customFPDestinationSOSWKST.Cells[10, 3].value = laborPlanInformation[4];//nights hours
                        customFPDestinationSOSWKST.Cells[11, 3].value = laborPlanInformation[5];//nights tph
                        customFPDestinationSOSWKST.Cells[9, 3].value = laborPlanInformation[6];//nights shipped
                        customFPDestinationSOSWKST.Cells[12, 3].value = laborPlanInformation[7];//nights ordered
                        break;
                    default:
                        customFPDestinationSOSWKST.Cells[5, 9].value = "Days";
                        break;
                }
                log.Info("Labor Plan information Updated in start of shift");
            }
            else if (laborPlanPop)
            {
                log.Fatal("{0} Labor Plan Information skipped due to lack of Plan Info", warehouse);
            }
            else log.Warn("{0} Labor Plan Information skipped via Populate flag", warehouse);

            //Hourly TPH Configuration
            for (int i = 0; i < 10; i++)
            {
                customFlowPlanDestinationHOURLYTPHWKST.Cells[25, i + 4].value = tphGoals[i];
            }
            log.Info("TPH assumptions Adjusted for AVP1");

            //Master Data Configuration
            customFlowPlanDestinationMasterDataWKST.Cells[30, 8].value = distrobutionList;
            log.Info("Master Data distrobution list updated to {0}", distrobutionList);
            customFlowPlanDestinationMasterDataWKST.Cells[12, 14].value = backlogHandoffPercent;
            log.Info("Master Data backlog handoff percentage is updated to {0}", backlogHandoffPercent);
            customFlowPlanDestinationMasterDataWKST.Cells[18, 17].value = VCPUWageRate;
            log.Info("Master Data VCPU wage rate is updated to {0}", VCPUWageRate);
            customFlowPlanDestinationMasterDataWKST.Cells[17, 14].value = DeadMan;
            log.Info("Master Data Dead Man is updated to {0}", DeadMan);


            log.Info("Master Data sections update completed");

            //Hourly ST 
            //tHis section is where we are going to read the current set values when we archive to allow the team to update the values and they will stick this means that i need to create and XML representation of each warehouse to store the defaults on both nights and days
            //I may want to extrapulate and create WareHouse objects so that i can keep track of all of these different setups and make adding warehouse extensible and flexible on runtime. this will reduce errors when the labor plans change.
            if(shift == "Days"){ daysRates.CopyTo(rates, 0);}
            else { nightsRate.CopyTo(rates, 0); }
            for (int i = 0; i <=8; i++)
            {
                customFlowPlanDestinationHourlySTWKST.Cells[i + 52, 4].value = rates[i];
            }
            
            //reset caluclation
            CustomiseFlowPlanApplication.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

                //save copy of file off some where
                string saveFilename = newFileDirectory 
                    + warehouse
                    + " "
                    + shift
                    + " Flow Plan V " + customFPDestinationSOSWKST.Cells[2,9].value  
                    + " " 
                    + System.DateTime.Now.ToString("yyyy-MM-dd");

                //Save as the workbook
                var fi = new FileInfo(saveFilename + ".xlsm");
                if (fi.Exists)
                {
                    try
                    {
                        File.Delete(saveFilename + ".xlsm");
                        log.Info("Existing file found and removed");
                    }
                    catch (IOException)
                    {
                    log.Fatal("File {0} unable to be removed, currently in use.", saveFilename + ".xlsm");
                    }
                }

                try
                {
                    customFlowPlanDestinationWB.SaveAs(saveFilename + ".xlsm");
                    customFlowPlanDestinationWB.Close();
                    log.Info("File Saved successfully at: {0}", saveFilename + ".xlsm");
                }
                catch (Exception)
                {
                    try
                    {
                        customFlowPlanDestinationWB.SaveAs(saveFilename + "(Empty).xlsm");
                        customFlowPlanDestinationWB.Close();
                        log.Warn("Origial file unable to be saved successfully, secondary save successful at  at: {0}", saveFilename + "(Empty).xlsm");
                    }
                    catch(Exception e)
                    {
                    log.Fatal("Both File attempts unable to be saved at: {0}", saveFilename + ".xlsm");
                    log.Fatal(e);
                    }
                }

            //close down the rest of the outstanding excel application.

            CustomiseFlowPlanApplication.Quit();
            customFlowPlanDestinationCHRGDATAWKST = null;
            customFlowPlanDestinationHOURLYTPHWKST = null;
            customFPDestinationSOSWKST = null;
            customFlowPlanDestinationWB = null;
            customizeFlowPlanWorkBookCollection = null; 
            chargeDataSourceWKST = null;
            chargeDataSourceWB = null;
            CustomiseFlowPlanApplication = null;

            //clear com object references
            if (customFlowPlanDestinationWB != null)
            { Marshal.Marshal.ReleaseComObject(customFlowPlanDestinationWB); }
            if (customFlowPlanDestinationCHRGDATAWKST != null)
            { Marshal.Marshal.ReleaseComObject(customFlowPlanDestinationCHRGDATAWKST); }
            if (chargeDataSourceWB != null)
            { Marshal.Marshal.ReleaseComObject(chargeDataSourceWB); }
            if (chargeDataSourceWKST != null)
            { Marshal.Marshal.ReleaseComObject(chargeDataSourceWKST); }
            if (customizeFlowPlanWorkBookCollection != null)
            { Marshal.Marshal.ReleaseComObject(customizeFlowPlanWorkBookCollection); }
            if (CustomiseFlowPlanApplication != null)
            { Marshal.Marshal.ReleaseComObject(CustomiseFlowPlanApplication); }


        }

        public static void GracefulExit(List<Warehouse> warehouseList)
        {
            try
            {
                File.WriteAllText(whXmlList, buildXmlFile(warehouseList));
                log.Info("Warehouse list written to XML file for future use.");
            }catch(Exception e)
            {
                log.Fatal(e);
            }
            
            LogManager.Flush();
        }
        public static void CrashshutDownApplication(Exception e)
        {
            log.Fatal("Crash shutdown Initiated");
            log.Fatal(e, "Causation of Abort:");
            LogManager.Flush();

        }

        private static string buildXmlFile (List<Warehouse> warehousesList)
        {
            XmlDocument warehousesXML = new XmlDocument();

            XmlNode warehouseRoot = warehousesXML.CreateElement("Warehouses");
            warehousesXML.AppendChild(warehouseRoot);

            foreach(Warehouse wh in warehousesList)
            {
                XmlNode warehouse = warehousesXML.CreateElement(wh.Name);
                warehouseRoot.AppendChild(warehouse);
                CreateNewChildXmlNode(warehousesXML, warehouse, "Name", wh.Name.ToString());
                CreateNewChildXmlNode(warehousesXML, warehouse, "Location", wh.Location.ToString());
                CreateNewChildXmlNode(warehousesXML, warehouse, "BlankLoc", wh.blankCopyLoc.ToString());
                CreateNewChildXmlNode(warehousesXML, warehouse, "ArchiveLoc", wh.archiveLoc.ToString());
                CreateNewChildXmlNode(warehousesXML, warehouse, "DistroListTarget", wh.DistroList.ToString());
                CreateNewChildXmlNode(warehousesXML, warehouse, "HandoffPercent", wh.HandoffPercent.ToString());
                CreateNewChildXmlNode(warehousesXML, warehouse, "VCPUWageRate", wh.VCPUWageRate.ToString());
                CreateNewChildXmlNode(warehousesXML, warehouse, "LaborPlanRows", string.Join(",", wh.laborPlanInforRows));
                CreateNewChildXmlNode(warehousesXML, warehouse, "TPHDistro", string.Join(",", wh.tphDistro));
                CreateNewChildXmlNode(warehousesXML, warehouse, "DaysRates", string.Join(",", wh.daysRate));
                CreateNewChildXmlNode(warehousesXML, warehouse, "NightsRates", string.Join(",", wh.nightsRate));
                CreateNewChildXmlNode(warehousesXML, warehouse, "PreShiftFlag", wh.PreShiftFlag.ToString());
                CreateNewChildXmlNode(warehousesXML, warehouse, "DeadMan", wh.DeadMan.ToString());
                CreateNewChildXmlNode(warehousesXML, warehouse, "LaborInfoPop", wh.laborInfoPop.ToString());
                CreateNewChildXmlNode(warehousesXML, warehouse, "PreshiftInforPop", wh.preshiftInfoPop.ToString());
                CreateNewChildXmlNode(warehousesXML, warehouse, "DaysChargePattern", string.Join(",", wh.daysChargePattern));
                CreateNewChildXmlNode(warehousesXML, warehouse, "NightsChargePattern", string.Join(",", wh.nightsChargePattern));
            }

            return warehousesXML.InnerXml;
        }

        public static void CreateNewChildXmlNode(XmlDocument document, XmlNode parentNode, string elementName, object value)
        {
            XmlNode node = document.CreateElement(elementName);
            node.AppendChild(document.CreateTextNode(value.ToString()));
            parentNode.AppendChild(node);
        }

        public static List<Warehouse> getWHList()
        {
            List<Warehouse> whList = new List<Warehouse>();
            int[] dprows = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            double[] rates = { 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            try
            {
                XmlDocument warehouseOpen = new XmlDocument();
                warehouseOpen.Load(whXmlList);
                Warehouse temp = new Warehouse("", "","",dprows, rates, rates, defaultTPHHourlyGoalDays );

                foreach (XmlNode warehouse in warehouseOpen.DocumentElement.ChildNodes)
                {
                    temp = new Warehouse("", "", "", dprows, rates, rates, defaultTPHHourlyGoalDays);
                    foreach (XmlNode warehouseinfo in warehouse.ChildNodes)
                    {
                        switch (warehouseinfo.Name)
                        {
                            case "Name":
                                temp.Name = warehouseinfo.InnerText;
                                break;
                            case "Location":
                                temp.Location = warehouseinfo.InnerText;
                                break;
                            case "BlankLoc":
                                temp.blankCopyLoc = warehouseinfo.InnerText;
                                break;
                            case "ArchiveLoc":
                                temp.archiveLoc = warehouseinfo.InnerText;
                                break;
                            case "DistroListTarget":
                                temp.DistroList = warehouseinfo.InnerText;
                                break;
                            case "HandoffPercent":
                                temp.HandoffPercent = double.Parse(warehouseinfo.InnerText);
                                break;
                            case "VCPUWageRate":
                                temp.VCPUWageRate = double.Parse(warehouseinfo.InnerText);
                                break;
                            case "LaborPlanRows":
                                temp.laborPlanInforRows = Array.ConvertAll(warehouseinfo.InnerText.Split(','), int.Parse);
                                break;
                            case "TPHDistro":
                                temp.tphDistro = Array.ConvertAll(warehouseinfo.InnerText.Split(','), double.Parse);
                                break;
                            case "DaysRates":
                                temp.daysRate = Array.ConvertAll(warehouseinfo.InnerText.Split(','), double.Parse);
                                break;
                            case "NightsRates":
                                temp.nightsRate = Array.ConvertAll(warehouseinfo.InnerText.Split(','), double.Parse);
                                break;
                            case "PreShiftFlag":
                                temp.PreShiftFlag = bool.Parse(warehouseinfo.InnerText);
                                break;
                            case "DeadMan":
                                temp.DeadMan = int.Parse(warehouseinfo.InnerText);
                                break;
                            case "LaborInfoPop":
                                temp.laborInfoPop = bool.Parse(warehouseinfo.InnerText);
                                break;
                            case "PreshiftInforPop":
                                temp.preshiftInfoPop = bool.Parse(warehouseinfo.InnerText);
                                break;
                            case "DaysChargePattern":
                                temp.daysChargePattern = Array.ConvertAll(warehouseinfo.InnerText.Split(','), double.Parse);
                                break;
                            case "NightsChargePattern":
                                temp.nightsChargePattern = Array.ConvertAll(warehouseinfo.InnerText.Split(','), double.Parse);
                                break;
                            default:
                                log.Warn("XML Node not found: {0}", warehouseinfo.Name);
                                break;
                        }
                    }

                    whList.Add(temp);
                    log.Info("{0} Warehouse added to list of warehouses", temp.Name);
                }
            }
            catch
            {
                log.Warn("Issues reading XML warehouse list reverting to default warehouse assumptions");
                //int[] dprows = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                int[] cfc1dprows = { 32, 35, 38, 42, 73, 76, 79, 83, 32, 35, 38, 42 };
                int[] avp1dprows = { 25, 28, 31, 35, 59, 62, 65, 69, 25, 28, 31, 35 };
                int[] dfw1dprows = { 34, 37, 40, 44, 74, 77, 80, 84, 34, 37, 40, 44 };
                int[] wfc2dprows = { 30, 33, 36, 40, 67, 70, 73, 77, 30, 33, 36, 40 };
                int[] efc3dprows = { 26, 29, 32, 36, 61, 64, 67, 71, 25, 29, 32, 36 };
                double[] daysRate = { };
                double[] nightsRate = { };
                double[] defaultGoalRates = { 47, 100, 24, 30, 110, 90, 40, 300, 2.5 };


                Warehouse AVP1 = new Warehouse("AVP1", @"\\avp1afs01\Outbound\AVP Flow Plan\Blank Copy\", @"\\avp1afs01\Outbound\AVP Flow Plan\FlowPlanArchive\NotProcessed\",avp1dprows, avp1TPHHourlyGoalDays, avp1TPHHourlyGoalNights, defaultGoalRates, "Wilkes-Barre, PA", "DL-AVP1-Outbound");
                Warehouse CFC1 = new Warehouse("CFC1", @"\\cfc1afs01\Outbound\Master Wash\Blank Copy\", @"\\cfc1afs01\Outbound\Master Wash\FlowPlanArchive\NotProcessed\", cfc1dprows, avp1TPHHourlyGoalDays, avp1TPHHourlyGoalNights, defaultGoalRates, "Clayton, IN", "DL-CFC-Outbound");
                Warehouse DFW1 = new Warehouse("DFW1", @"\\dfw1afs01\Outbound\Blank Flow Plan\", @"\\dfw1afs01\Outbound\FlowPlanArchive\NotProcessed\", dfw1dprows, avp1TPHHourlyGoalDays, defaultGoalRates, defaultGoalRates, "Dallas, Texas", "DL-DFW1-Outbound");
                Warehouse EFC3 = new Warehouse("EFC3", @"\\wh-pa-fs-01\OperationsDrive\Blank Flow Plan\", @"\\wh-pa-fs-01\OperationsDrive\FlowPlanArchive\NotProcessed\", efc3dprows, defaultTPHHourlyGoalDays, defaultGoalRates, defaultGoalRates, "Mechanicsburg, PA", "DL-EFC3-Outbound");
                Warehouse WFC2 = new Warehouse("WFC2", @"\\wfc2afs01\Outbound\Outbound Flow Planner\Blank Copy\", @"\\wfc2afs01\Outbound\Outbound Flow Planner\FlowPlanArchive\NotProcessed\",  wfc2dprows, defaultTPHHourlyGoalDays, defaultGoalRates, defaultGoalRates, "Reno, NV", "DL-WFC2-Outbound");


                whList.Add(AVP1);
                whList.Add(CFC1);
                whList.Add(DFW1);
                whList.Add(EFC3);
                whList.Add(WFC2);
            }



            return whList;
        }

        public static void CustomizePreShift (double[] laborPlanInfo, string shift, string blankLocation, string warehouse, bool infoPop)//need to enhance logging and pull calculations out of this function.
        {

            //define Excel Application for the Creation Process
            Excel.Application CustomisePreShiftApplication = null;
            CustomisePreShiftApplication = new Excel.Application();
            CustomisePreShiftApplication.Visible = visibilityFlag;
            CustomisePreShiftApplication.DisplayAlerts = alertsFlag;

            //create the empty containers for the excel application
            Excel.Workbooks customizePreShiftWorkBookCollection = null;
            Excel.Workbook customSupPreShiftDestinationWB = null;
            Excel.Worksheet SupINDWKST = null;
            Excel.Worksheet SupPACKWKST = null;
            Excel.Worksheet SupPICKWKST = null;
            Excel.Worksheet SupDockWKST = null;

            Excel.Workbook customLeadPreShiftDestinationWB = null;
            Excel.Worksheet LeadINDWKST = null;
            Excel.Worksheet LeadPACKWKST = null;
            Excel.Worksheet LeadPICKWKST = null;
            Excel.Worksheet LeadDockWKST = null;

            log.Info("Excel Empty containers created for Pre-Shift Documentation");

            customizePreShiftWorkBookCollection = CustomisePreShiftApplication.Workbooks;
            try
            {
                log.Info("Opening Master Sup Preshift at {0}", masterPreShiftCheckList);
                customSupPreShiftDestinationWB = customizePreShiftWorkBookCollection.Open(masterPreShiftCheckList + "Standard Work - OB Supervisors(MCRC).xlsx", false, false);
            }catch
            {
                log.Fatal("Could not open master Sup Preshift, skipping {0}' customization", shift);
                return;
            }

            try
            {
                log.Info("Opening Master Lead Preshift at {0}", masterPreShiftCheckList);
                customLeadPreShiftDestinationWB = customizePreShiftWorkBookCollection.Open(masterPreShiftCheckList + "Standard Work - OB Leads(MCRC).xlsx", false, false);
            }
            catch
            {
                log.Fatal("Could not open master Lead Preshift, skipping {0}' customization", shift);
                return;
            }

            SupINDWKST = customSupPreShiftDestinationWB.Worksheets.Item["Induction Sup"];
            SupPACKWKST = customSupPreShiftDestinationWB.Worksheets.Item["Pack Sup"];
            SupPICKWKST = customSupPreShiftDestinationWB.Worksheets.Item["Pick Sup"];
            SupDockWKST = customSupPreShiftDestinationWB.Worksheets.Item["Dock and VNA Sup"];

            LeadINDWKST = customLeadPreShiftDestinationWB.Worksheets.Item["Induction Lead"];
            LeadPACKWKST = customLeadPreShiftDestinationWB.Worksheets.Item["Pack Lead"];
            LeadPICKWKST = customLeadPreShiftDestinationWB.Worksheets.Item["Pick Lead"];
            LeadDockWKST = customLeadPreShiftDestinationWB.Worksheets.Item["Dock Lead"];

            if(infoPop)
            {
                if (shift == "Days")//should really search for the Item names in the event that we make changes i can still find where to put the info...
                {
                    SupINDWKST.Cells[13, 7].value = Math.Round(laborPlanInfo[2]);
                    SupINDWKST.Cells[14, 7].value = laborPlanInfo[1];
                    SupINDWKST.Cells[4, 3].value = laborPlanInfo[1];

                    SupPACKWKST.Cells[13, 7].value = Math.Round(laborPlanInfo[2]);
                    SupPACKWKST.Cells[14, 7].value = laborPlanInfo[1];
                    SupPACKWKST.Cells[4, 3].value = calcTotalSinglesShip(laborPlanInfo[2]);
                    SupPACKWKST.Cells[5, 3].value = calcSinglesHourlyShip(laborPlanInfo[2]);

                    SupPICKWKST.Cells[13, 7].value = Math.Round(laborPlanInfo[2]);
                    SupPICKWKST.Cells[14, 7].value = laborPlanInfo[1];
                    SupPICKWKST.Cells[4, 3].value = laborPlanInfo[1];
                    SupPICKWKST.Cells[5, 3].value = calcTotalHourlyShip(laborPlanInfo[2]);
                    SupPICKWKST.Cells[6, 3].value = calcTotalHourlyShip(laborPlanInfo[2]) * breakModifier;
                    SupPICKWKST.Cells[7, 3].value = calcMultiPickers(calcMultisHourlyShip(laborPlanInfo[2]));
                    SupPICKWKST.Cells[8, 3].value = calcTotalMultiShip(laborPlanInfo[2]);
                    SupPICKWKST.Cells[9, 3].value = calcMultisHourlyShip(laborPlanInfo[2]);
                    SupPICKWKST.Cells[10, 3].value = calcMultisHourlyShip(laborPlanInfo[2]) * breakModifier;

                    SupDockWKST.Cells[13, 7].value = Math.Round(laborPlanInfo[2]);
                    SupDockWKST.Cells[14, 7].value = laborPlanInfo[1];
                    SupDockWKST.Cells[4, 3].value = calcDocksHourlycont(laborPlanInfo[2]);

                    LeadINDWKST.Cells[13, 7].value = Math.Round(laborPlanInfo[2]);
                    LeadINDWKST.Cells[14, 7].value = laborPlanInfo[1];
                    LeadINDWKST.Cells[4, 3].value = laborPlanInfo[1];

                    LeadPACKWKST.Cells[13, 7].value = Math.Round(laborPlanInfo[2]);
                    LeadPACKWKST.Cells[14, 7].value = laborPlanInfo[1];
                    LeadPACKWKST.Cells[4, 3].value = calcTotalSinglesShip(laborPlanInfo[2]);
                    LeadPACKWKST.Cells[5, 3].value = calcSinglesHourlyShip(laborPlanInfo[2]);

                    LeadPICKWKST.Cells[13, 7].value = Math.Round(laborPlanInfo[2]);
                    LeadPICKWKST.Cells[14, 7].value = laborPlanInfo[1];
                    LeadPICKWKST.Cells[4, 3].value = laborPlanInfo[1];
                    LeadPICKWKST.Cells[5, 3].value = calcTotalHourlyShip(laborPlanInfo[2]);
                    LeadPICKWKST.Cells[6, 3].value = calcTotalHourlyShip(laborPlanInfo[2]) * breakModifier;
                    LeadPICKWKST.Cells[7, 3].value = calcMultiPickers(calcMultisHourlyShip(laborPlanInfo[2]));
                    LeadPICKWKST.Cells[8, 3].value = calcTotalMultiShip(laborPlanInfo[2]);
                    LeadPICKWKST.Cells[9, 3].value = calcMultisHourlyShip(laborPlanInfo[2]);
                    LeadPICKWKST.Cells[10, 3].value = calcMultisHourlyShip(laborPlanInfo[2]) * breakModifier;

                    LeadDockWKST.Cells[13, 7].value = Math.Round(laborPlanInfo[2]);
                    LeadDockWKST.Cells[14, 7].value = laborPlanInfo[1];
                    LeadDockWKST.Cells[4, 3].value = calcDocksHourlycont(laborPlanInfo[2]);
                }
                else
                {
                    SupINDWKST.Cells[13, 7].value = Math.Round(laborPlanInfo[6]);
                    SupINDWKST.Cells[14, 7].value = laborPlanInfo[5];
                    SupINDWKST.Cells[4, 3].value = laborPlanInfo[5];

                    SupPACKWKST.Cells[13, 7].value = Math.Round(laborPlanInfo[6]);
                    SupPACKWKST.Cells[14, 7].value = laborPlanInfo[5];
                    SupPACKWKST.Cells[4, 3].value = calcTotalSinglesShip(laborPlanInfo[6]);
                    SupPACKWKST.Cells[5, 3].value = calcSinglesHourlyShip(laborPlanInfo[6]);

                    SupPICKWKST.Cells[13, 7].value = Math.Round(laborPlanInfo[6]);
                    SupPICKWKST.Cells[14, 7].value = laborPlanInfo[5];
                    SupPICKWKST.Cells[4, 3].value = Math.Round(laborPlanInfo[6]);
                    SupPICKWKST.Cells[5, 3].value = calcTotalHourlyShip(laborPlanInfo[6]);
                    SupPICKWKST.Cells[6, 3].value = calcTotalHourlyShip(laborPlanInfo[6]) * breakModifier;
                    SupPICKWKST.Cells[7, 3].value = calcMultiPickers(calcMultisHourlyShip(laborPlanInfo[6]));
                    SupPICKWKST.Cells[8, 3].value = calcTotalMultiShip(laborPlanInfo[6]);
                    SupPICKWKST.Cells[9, 3].value = calcMultisHourlyShip(laborPlanInfo[6]);
                    SupPICKWKST.Cells[10, 3].value = calcMultisHourlyShip(laborPlanInfo[6]) * breakModifier;

                    SupDockWKST.Cells[13, 7].value = Math.Round(laborPlanInfo[6]);
                    SupDockWKST.Cells[14, 7].value = laborPlanInfo[5];
                    SupDockWKST.Cells[4, 3].value = calcDocksHourlycont(laborPlanInfo[6]);

                    LeadINDWKST.Cells[13, 7].value = Math.Round(laborPlanInfo[6]);
                    LeadINDWKST.Cells[14, 7].value = laborPlanInfo[5];
                    LeadINDWKST.Cells[4, 3].value = laborPlanInfo[5];

                    LeadPACKWKST.Cells[13, 7].value = Math.Round(laborPlanInfo[6]);
                    LeadPACKWKST.Cells[14, 7].value = laborPlanInfo[5];
                    LeadPACKWKST.Cells[4, 3].value = calcTotalSinglesShip(laborPlanInfo[6]);
                    LeadPACKWKST.Cells[5, 3].value = calcSinglesHourlyShip(laborPlanInfo[6]);

                    LeadPICKWKST.Cells[13, 7].value = Math.Round(laborPlanInfo[6]);
                    LeadPICKWKST.Cells[14, 7].value = laborPlanInfo[5];
                    LeadPICKWKST.Cells[4, 3].value = Math.Round(laborPlanInfo[6]);
                    LeadPICKWKST.Cells[5, 3].value = calcTotalHourlyShip(laborPlanInfo[6]);
                    LeadPICKWKST.Cells[6, 3].value = calcTotalHourlyShip(laborPlanInfo[6]) * breakModifier;
                    LeadPICKWKST.Cells[7, 3].value = calcMultiPickers(calcMultisHourlyShip(laborPlanInfo[6]));
                    LeadPICKWKST.Cells[8, 3].value = calcTotalMultiShip(laborPlanInfo[6]);
                    LeadPICKWKST.Cells[9, 3].value = calcMultisHourlyShip(laborPlanInfo[6]);
                    LeadPICKWKST.Cells[10, 3].value = calcMultisHourlyShip(laborPlanInfo[6]) * breakModifier;

                    LeadDockWKST.Cells[13, 7].value = Math.Round(laborPlanInfo[6]);
                    LeadDockWKST.Cells[14, 7].value = laborPlanInfo[5];
                    LeadDockWKST.Cells[4, 3].value = calcDocksHourlycont(laborPlanInfo[6]);
                }
            }
            
            if (!Directory.Exists(blankLocation + @"\Preshifts\"))
                Directory.CreateDirectory(blankLocation + @"\Preshifts\");
            string saveFileName = blankLocation + @"\Preshifts\" +
                warehouse + " "+
                shift;
            try
            {
                customSupPreShiftDestinationWB.SaveAs(saveFileName + " Sup Pre-Shift Check List.xlsx");
                log.Info("{0} {1} Preshift Saved at {2}", warehouse, shift, saveFileName + " Sup Pre-Shift Check List.xlsx");
            }
            catch
            {
                try
                {
                    log.Warn("Issue saving {0} Sup Preshift, trying secondary save function.", shift);
                    customSupPreShiftDestinationWB.SaveAs(saveFileName + " Sup Pre-Shift Check List(new).xlsx");
                }
                catch
                {
                    log.Fatal("Unable to save {0} Sup Preshift skipping....", shift);
                }
            }
            try
            {
                customLeadPreShiftDestinationWB.SaveAs(saveFileName + " Lead Pre-Shift Check List.xlsx");
                log.Info("{0} {1} Preshift Saved at {2}", warehouse, shift, saveFileName + " Lead Pre-Shift Check List.xlsx");
            }
            catch
            {
                try
                {
                    log.Warn("Issue saving {0} Lead Preshift, trying secondary save function.", shift);
                    customLeadPreShiftDestinationWB.SaveAs(saveFileName + " Lead Pre-Shift Check List(new).xlsx");
                }
                catch
                {
                    log.Fatal("Unable to save {0} Lead Preshift skipping....", shift);
                }
             }
           

            customizePreShiftWorkBookCollection = null;
            customSupPreShiftDestinationWB = null;
            customLeadPreShiftDestinationWB = null;
            SupINDWKST = null;
            SupPACKWKST = null;
            SupPICKWKST = null;
            SupDockWKST = null;
            LeadINDWKST = null;
            LeadPACKWKST = null;
            LeadPICKWKST = null;
            LeadDockWKST = null;
            CustomisePreShiftApplication.Quit();
            CustomisePreShiftApplication = null;

        }
        //public static void gatherRawChargeDate
        //{

        //    foreach(row in data)
        //        {
        //        switch col1
        //    {
        //            case "AVP1":
        //                foreach (hour in day)
        //                {
        //                    avp1source[DateTime, hour] = avp1TPHHourlyGoalDays;

        //                }
        //                break;
        //        }
        //    }
                
        //}
        public static double calcTotalHourlyShip(double shipGoal)
        {

            return Math.Round(shipGoal / shiftLength);
        }

        public static double calcTotalSinglesShip(double shipGoal)
        {

            return Math.Round(shipGoal * singlesPercentSplit);
        }

        public static double calcSinglesHourlyShip(double shipGoal)
        {

            return Math.Round(shipGoal / shiftLength * singlesPercentSplit);
        }

        public static double calcTotalMultiShip(double shipGoal)
        {

            return Math.Round(shipGoal * (1-singlesPercentSplit));
        }

        public static double calcMultisHourlyShip(double shipGoal)
        {

            return Math.Round(shipGoal / shiftLength * (1-singlesPercentSplit));
        }

        public static double calcDocksHourlycont(double shipGoal)
        {

            return Math.Round(shipGoal/ shiftLength / UPC);
        }
        
        public static double calcMultiPickers(double shipGoal)
        {
            return Math.Round(shipGoal / defaultMultiRate);
        }

        public static void GatherSQLData()
        {

            using (SqlConnection connection =
                   new SqlConnection(prodConnectionString))
            {
                SqlCommand command =
                    new SqlCommand(chargeQuery, connection);
                connection.Open();

                SqlDataReader reader = command.ExecuteReader();

                Console.WriteLine("{0}, {1}, {2}, {3}, {4}", reader.GetName(0), reader.GetName(1), reader.GetName(2), reader.GetName(3), reader.GetName(4));
                for (int x = 0; x < 5; x++)
                {
                    
                }
                int y = 2;
                // Call Read before accessing data.
                while (reader.Read())
                {
                   
                }

                // Call Close when done reading.
                reader.Close();
            }
        }
}

       



    




