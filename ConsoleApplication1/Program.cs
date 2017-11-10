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

namespace FlowPlanConstruction
{
    class Program
    {
        private static Logger log = LogManager.GetCurrentClassLogger();

        //section toggles
        private const bool debugMode = false;
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
        public static double[] avp1TPHHourlyGoalNights= { 0.80, 1.12, 0.80, 1.12, 1.08, 0.96, 1.12, 0.80, 1.12, 1.04 };
        public static double[] defaultTPHHourlyGoalDays = { .90, 1, .90,1,1,.7,.9, 1,.9, 1,.7};
        private static double defaultBackLogHandoff = .46;
        public static double singlesPercentSplit = .33;
        public static double shiftLength = 9.5;
        public static double breakModifier = .7;
        public static double UPC = 2.5;
        public static double defaultMultiRate = 47;

        //add a way to edit the XML file external to this program could be stand alone configuration that would allow user to kick off a run of the constructor with new perimeters....?


        static void Main(string[] args)
        {
            //setup log file information
            string user = WindowsIdentity.GetCurrent().Name;

            log.Info("Application started by {0}", user);

            bool runLaborPlanPopulate = false;

            double[] laborplaninfo = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

            List<Warehouse> warehouseList = new List<Warehouse>();

            warehouseList = getWHList();

            foreach( Warehouse wh in warehouseList)
            {
                if(!debugMode && wh.Name =="AVP1")
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
                        CleanUpDirectories(wh.blankCopyLoc, wh.archiveLoc);
                    }
                        
                    if(runCustom)
                    {
                        log.Info("Flow Plan Customization process Initalized for {0}", wh.Name);
                        foreach (string shift in shifts)
                        {
                            customizeFlowPlan(wh.blankCopyLoc, wh.Name, shift, runLaborPlanPopulate, laborplaninfo, wh.DistroList, wh.HandoffPercent, wh.VCPUWageRate);
                            if(wh.Name == "AVP1" || wh.Name == "CFC1")
                                CustomizePreShift(laborplaninfo, shift, wh.blankCopyLoc, wh.Name);
                        }
                    }
                }
            }

            GracefulExit(warehouseList);
            
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

        private static void customizeFlowPlan(string newFileDirectory, string warehouse, string shift,bool laborPlanInfoAvaliable, double[] laborPlanInformation, string distrobutionList, double backlogHandoffPercent, double VCPUWageRate)// this may need restructuring to make it more readable keep and eye on for future improvements
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

            //object test = chargeDataSourceWKST.UsedRange.Value;

            

            log.Info("New charge Data inserted into the master copy");

            chargeDataSourceWB.Close();

            //setup SOS Information
            customFPDestinationSOSWKST.Cells[3, 9].value = warehouse;
            customFPDestinationSOSWKST.Cells[4, 9].value = System.DateTime.Now.ToString("yyyy-MM-dd");
            customFPDestinationSOSWKST.Cells[5, 9].value = shift;

            log.Info("Start of shift information updated: Warehouse: {0} Date: {1} Shift: {2}", warehouse, System.DateTime.Now.ToString("yyyy-MM-dd"), shift);

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
                        customFPDestinationSOSWKST.Cells[5, 9].value = "Days";
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
            log.Info("Master Data distrobution list updated to {0}", distrobutionList);
            customFlowPlanDestinationMasterDataWKST.Cells[12, 14].value = backlogHandoffPercent;
            log.Info("Master Data backlog handoff percentage is updated to {0}", backlogHandoffPercent);
            customFlowPlanDestinationMasterDataWKST.Cells[18, 17].value = VCPUWageRate;
            log.Info("Master Data VCPU wage rate is updated to {0}", VCPUWageRate);

            log.Info("Master data sections update completed");

            //Hourly ST 
            //tHis section is where we are going to read the current set values when we archive to allow the team to update the values and they will stick this means that i need to create and XML representation of each warehouse to store the defaults on both nights and days
            //I may want to extrapulate and create WareHouse objects so that i can keep track of all of these different setups and make adding warehouse extensible and flexible on runtime. this will reduce errors when the labor plans change.

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
        public void CrashshutDownApplication(Exception e)
        {

            log.Fatal(e, "Causation of Abort:");

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
            try
            {
                XmlDocument warehouseOpen = new XmlDocument();
                warehouseOpen.Load(whXmlList);
                Warehouse temp = new Warehouse("", "", "", "", "", dprows, defaultTPHHourlyGoalDays, defaultTPHHourlyGoalDays, defaultTPHHourlyGoalDays);

                foreach (XmlNode warehouse in warehouseOpen.DocumentElement.ChildNodes)
                {
                    temp = new Warehouse("", "", "", "", "", dprows, defaultTPHHourlyGoalDays, defaultTPHHourlyGoalDays, defaultTPHHourlyGoalDays);
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

                

                Warehouse AVP1 = new Warehouse("AVP1", "Scranton/Wilkes-barre, PA", @"\\avp1afs01\Outbound\AVP Flow Plan\Blank Copy\", @"\\avp1afs01\Outbound\AVP Flow Plan\FlowPlanArchive\NotProcessed\", "DL-AVP1-Outbound@chewy.com", avp1dprows, avp1TPHHourlyGoalDays, defaultGoalRates, defaultGoalRates);
                Warehouse CFC1 = new Warehouse("CFC1", "Clayton, IN", @"\\cfc1afs01\Outbound\Master Wash\Blank Copy\", @"\\cfc1afs01\Outbound\Master Wash\FlowPlanArchive\NotProcessed\", "DL-CFC-Outbound@chewy.com", cfc1dprows, defaultTPHHourlyGoalDays, defaultGoalRates, defaultGoalRates);
                Warehouse DFW1 = new Warehouse("DFW1", "Dallas, TX", @"\\dfw1afs01\Outbound\Blank Flow Plan\", @"\\dfw1afs01\Outbound\FlowPlanArchive\NotProcessed\", "DL-DFW1-Outbound@chewy.com", dfw1dprows, avp1TPHHourlyGoalDays, defaultGoalRates, defaultGoalRates);
                Warehouse EFC3 = new Warehouse("EFC3", "Mechanicsburg, PA", @"\\wh-pa-fs-01\OperationsDrive\Blank Flow Plan\", @"\\wh-pa-fs-01\OperationsDrive\FlowPlanArchive\NotProcessed\", "DL-EFC3-Outbound@chewy.com", efc3dprows, defaultTPHHourlyGoalDays, defaultGoalRates, defaultGoalRates);
                Warehouse WFC2 = new Warehouse("WFC2", "Reno, NV", @"\\wfc2afs01\Outbound\Outbound Flow Planner\Blank Copy\", @"\\wfc2afs01\Outbound\Outbound Flow Planner\FlowPlanArchive\NotProcessed\", "DL-WFC2-Outbound@chewy.com", wfc2dprows, defaultTPHHourlyGoalDays, defaultGoalRates, defaultGoalRates);


                whList.Add(AVP1);
                whList.Add(CFC1);
                whList.Add(DFW1);
                whList.Add(EFC3);
                whList.Add(WFC2);
            }



            return whList;
        }

        public static void CustomizePreShift (double[] laborPlanInfo, string shift, string blankLocation, string warehouse)//need to enhance logging and pull calculations out of this function.
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
                customSupPreShiftDestinationWB = customizePreShiftWorkBookCollection.Open(masterPreShiftCheckList + "Standard Work - Supervisors(MCRC).xlsx", false, false);
            }catch
            {
                log.Fatal("Could not open master Preshift, skipping {0}' customization", shift);
                return;
            }

            try
            {
                log.Info("Opening Master Lead Preshift at {0}", masterPreShiftCheckList);
                customLeadPreShiftDestinationWB = customizePreShiftWorkBookCollection.Open(masterPreShiftCheckList + "Standard Work - Leads(MCRC).xlsx", false, false);
            }
            catch
            {
                log.Fatal("Could not open master Preshift, skipping {0}' customization", shift);
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
                    log.Warn("Issue saving {0}'s Sup Preshift, trying secondary save function.", shift);
                    customSupPreShiftDestinationWB.SaveAs(saveFileName + " Sup Pre-Shift Check List.new.xlsx");
                }
                catch
                {
                    log.Fatal("Unable to save {0}'s Sup Preshift skipping....");
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
                    log.Warn("Issue saving {0}'s Lead Preshift, trying secondary save function.", shift);
                    customLeadPreShiftDestinationWB.SaveAs(saveFileName + " Lead Pre-Shift Check List.new.xlsx");
                }
                catch
                {
                    log.Fatal("Unable to save {0}'s Lead Preshift skipping....");
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
    }
}

       



    




