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

namespace FlowPlanConstruction
{
    class Program
    {
        private static Logger log = LogManager.GetCurrentClassLogger();

        public const bool visibilityFlag = false;
        public const bool alertsFlag = false;
        public const bool runArchive = true;
        public const bool runLaborPlan = true;
        public const string chargeDataSrcPath = @"\\cfc1afs01\Operations-Analytics\RAW Data\chargepattern.csv";
        public const string masterFlowPlanLocation = @"\\CFC1AFS01\Operations-Analytics\Projects\Flow Plan\Outbound Flow Plan v2.0(MCRC).xlsm";
        public const string laborPlanLocationation = @"\\chewy.local\bi\BI Community Content\Finance\Labor Models\";
        public static string[] warehouses = { "AVP1", /*"CFC1", "DFW1", "EFC3", "WFC2" */};
        public static string[] shifts = { "Days", "Nights", "" };

        public static double[] avp1TPHHourlyGoalDays = { 0.80, 1.12, 1.12, 0.80, 1.08, 0.96, 1.12, 0.80, 1.12, 1.04 };
        public static double[] avp1TPHHourlyGoalNights= { 0.80, 1.12, 0.80, 1.12, 1.08, 0.96, 1.12, 0.80, 1.12, 1.04 };
        //public static double[] OtherFCTPHHourlyGoal = { 0.80, 1.12, 0.80, 1.12, 1.08, 0.96, 1.12, 0.80, 1.12, 1.04 };

        static void Main(string[] args)
        {
            
            //setup log file information
            string user = WindowsIdentity.GetCurrent().Name;

            log.Trace("Application started by {0}", user);

            string currentWarehouse = "";
            string distrobutionList = "";

            bool runLaborPlanPopulate = false;

            var createdFileDestination = @"\\CFC1AFS01\Operations-Analytics\Projects\Flow Plan\";

            //set up enviroment for each FC
            string archivePath = @"\\cfc1afs01\Operations-Analytics\Projects\Flow Plan\BackUpArchive";
            int[] dprows = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            int[] cfc1dprows = { 31, 34, 37, 41, 71, 74, 77, 81, 31, 34, 37, 41 };
            int[] avp1dprows = { 25, 28, 31, 35, 59, 62, 65, 69, 25, 28, 31, 35 };
            int[] dfw1dprows = { 27, 30, 33, 37, 60, 63, 66, 70, 27, 30, 33, 37 };
            int[] wfc2dprows = { 30, 33, 36, 40, 67, 70, 73, 77, 30, 33, 36, 40 };
            int[] efc3dprows = { 25, 28, 31, 35, 59, 62, 65, 69, 25, 28, 31, 35 };

            double[] laborplaninfo = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

            foreach (string wh in warehouses)
            {
                currentWarehouse = wh;
                switch (wh)
                {
                    case "AVP1":
                        createdFileDestination = @"\\avp1afs01\Outbound\AVP Flow Plan\Blank Copy\";
                        archivePath = @"\\avp1afs01\Outbound\AVP Flow Plan\FlowPlanArchive\NotProcessed\";
                        avp1dprows.CopyTo(dprows, 0);
                        distrobutionList = "DL-AVP1-Outbound@chewy.com";
                        break;
                    case "CFC1":
                        createdFileDestination = @"\\cfc1afs01\Outbound\Master Wash\Blank Copy\";
                        archivePath = @"\\cfc1afs01\Outbound\Master Wash\FlowPlanArchive\NotProcessed\";
                        cfc1dprows.CopyTo(dprows, 0);
                        distrobutionList = "DL-CFC-Outbound@chewy.com";
                        break;
                    case "DFW1":
                        createdFileDestination = @"\\dfw1afs01\Outbound\Blank Flow Plan\";
                        archivePath = @"\\dfw1afs01\Outbound\FlowPlanArchive\NotProcessed\";
                        dfw1dprows.CopyTo(dprows, 0);
                        distrobutionList = "DL-DFW1-Outbound@chewy.com";
                        break;
                    case "EFC3":
                        createdFileDestination = @"\\wh-pa-fs-01\OperationsDrive\Blank Flow Plan\";
                        archivePath = @"\\wh-pa-fs-01\OperationsDrive\FlowPlanArchive\NotProcessed\";
                        efc3dprows.CopyTo(dprows, 0);
                        distrobutionList = "DL-EFC3-Outbound@chewy.com";
                        break;
                    case "WFC2":
                        createdFileDestination = @"\\wfc2afs01\Outbound\Outbound Flow Planner\Blank Copy\";
                        archivePath = @"\\wfc2afs01\Outbound\Outbound Flow Planner\FlowPlanArchive\NotProcessed\";
                        wfc2dprows.CopyTo(dprows, 0);
                        distrobutionList = "DL-WFC2-Outbound@chewy.com";
                        break;
                    default:
                        log.Warn("Warehouse {0} not found please add to structure." , currentWarehouse);
                        break;
                }

                if (runLaborPlan)
                {
                    runLaborPlanPopulate = obtainLaborPlanInfo(dprows, wh, laborplaninfo);
                }

                if (runArchive)
                {
                    CleanUpDirectories(createdFileDestination, archivePath);
                }

                foreach (string shift in shifts)
                {
                    customizeFlowPlan(createdFileDestination, currentWarehouse, shift, runLaborPlan, laborplaninfo, distrobutionList);
                }
            }
        }
        private static void CleanUpDirectories(string locToArchive, string archiveLocation)
        {
            
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
                    log.Info("File Coppied from: " + copy + " to Archive");
                    System.IO.File.Copy(copy, archivePathFile, true);
                    System.IO.File.Delete(copy);
                }
                catch (System.IO.IOException)
                {
                    log.Warn(System.DateTime.Now + ":\t" + "File currently in use: " + copy);
                }
            }
        }

        private static bool obtainLaborPlanInfo( int[]dprows, string wh, double[] laborplaninfo)
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
                        return false;
                    }
                    
                    Excel.Worksheet OBDP = laborPlanModel.Worksheets.Item["OB Daily Plan"];

                    //determine which rows to go look for
                    foreach (string label in rowlabels)
                    {
                        
                        if (r > 11)
                        {
                            log.Warn("Current Row Check Greater than number of verifications needed.. Shutting down Labor Plan Search");
                            return false;
                        }
                        else
                        {


                            string rcheck = OBDP.Cells[dprows[r], 2].value; //Set the row check = to what we think the row should be
                            if (rcheck == label)
                            {
                                log.Info("Row location verified for " + label + " at row " + dprows[r]);
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
                                        log.Warn("Row location for " + label + " now found at row " + dprows[r]);
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
                            return false;
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

            return true; // we want to add the data to the flow planners
        }

        private static void customizeFlowPlan(string newFileDirectory, string warehouse, string shift,bool laborPlanInfoAvaliable, double[] laborPlanInformation, string distrobutionList)
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
            customFlowPlanDestinationMasterDataWKST = customFlowPlanDestinationWB.Worksheets.Item["Master Data"];

            log.Info("Charge Data Source: {0}", chargeDataSrcPath);
            log.Info("Master Flow Plan Source: {0}", masterFlowPlanLocation);

            //set application calculation method
            CustomiseFlowPlanApplication.Calculation = Excel.XlCalculation.xlCalculationManual;

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

            object test = chargeDataSourceWKST.UsedRange.Value;

            

            log.Info("New charge Data inserted into the master copy");

            chargeDataSourceWB.Close();

            //setup SOS Information
            customFPDestinationSOSWKST.Cells[3, 9].value = warehouse;
            customFPDestinationSOSWKST.Cells[4, 9].value = System.DateTime.Now.ToString("yyyy-MM-dd");
            customFPDestinationSOSWKST.Cells[5, 9].value = shift;

            if (laborPlanInfoAvaliable)
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
                        break;
                }
                log.Info("Labor Plan information Updated in start of shift");
            }
            else log.Warn("Labor Plan Information skipped");
                
                //Hourly TPH Configuration
                if (warehouse == "AVP1")
                {
                    if(shift =="Days")
                {
                    for (int i = 0; i < 10; i++)
                    {
                        customFlowPlanDestinationHOURLYTPHWKST.Cells[25, i + 4].value = avp1TPHHourlyGoalDays[i];
                    }
                }else
                {
                    for (int i = 0; i < 10; i++)
                    {
                        customFlowPlanDestinationHOURLYTPHWKST.Cells[25, i + 4].value = avp1TPHHourlyGoalNights[i];
                    }
                }

                    
                log.Info("TPH assumptions Adjusted for AVP1");
                }

            //Master Data Configuration
            customFlowPlanDestinationMasterDataWKST.Cells[30, 8].value = distrobutionList;

                //reset caluclation
                CustomiseFlowPlanApplication.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

                //save copy of file off some where
                string saveFilename = newFileDirectory 
                    + customFPDestinationSOSWKST.Cells[3, 9].value 
                    + " Flow Plan V " + customFPDestinationSOSWKST.Cells[2,9].value 
                    + " "
                    + shift 
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
                    log.Fatal("IOException thrown");
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
    }
}

       



    




