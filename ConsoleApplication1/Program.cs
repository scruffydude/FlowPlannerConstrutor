using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Marshal = System.Runtime.InteropServices;
using System.Security.Principal;

namespace FlowPlanConstruction
{
    class Program
    {
        
        static void Main(string[] args)
        {
            //setup log file information
            string logPath = @"\\cfc1afs01\Operations-Analytics\Log_Files\";
            StreamWriter logging = null;
            logging = new StreamWriter(logPath + "ChargePatternExecLog.txt");
            string user = WindowsIdentity.GetCurrent().Name;
            logging.WriteLine("Log Started for Charge Pattern run at " + System.DateTime.Now + " by " + user);
            Console.WriteLine("Log Started for Charge Pattern run at " + System.DateTime.Now + " by " + user);

            //Setup used Ranges variables
            int srcUsedRange = 0;
            int destUsedRange = 0;
            string filename = "";
            string a = "";
            bool runArchive = true;
            bool runLaborPlan = true;
            bool runLaborPlanPopulate = true;
            logging.WriteLine(System.DateTime.Now + ":\t" + "Success in Declaring Variables srcUsedRange & destUsedRange.");
            Console.WriteLine(System.DateTime.Now + ":\t" + "Success in Declaring Variables srcUsedRange & destUsedRange.");

            //Setup path variable
            string srcPath = @"\\cfc1afs01\Operations-Analytics\RAW Data\chargepattern.csv";
            string flowPlanMaster = @"\\CFC1AFS01\Operations-Analytics\Projects\Flow Plan\Outbound Flow Plan v2.0(MCRC).xlsm";
            var createdFileDestination = @"\\CFC1AFS01\Operations-Analytics\Projects\Flow Plan\";

            //setup Excel Application
            Excel.Application App = null;
            App = new Excel.Application();
            App.Visible = false;
            logging.WriteLine(System.DateTime.Now + ":\t" + "Success in setting up application Reference for Excel Application.");
            Console.WriteLine(System.DateTime.Now + ":\t" + "Success in setting up application Reference for Excel Application.");

            //create soure and destination files
            Excel.Workbooks Workbooks = null;
            Excel.Workbook srcWorkbook = null;
            Excel.Worksheet srcWorksheet = null;
            Excel.Workbook destWorkbook = null;
            Excel.Worksheet destWorksheet = null;

            //Assign the source and destination files
            Workbooks = App.Workbooks;
            srcWorkbook = Workbooks.Open(srcPath, false, false);
            srcWorksheet = srcWorkbook.Worksheets.Item[1];
            destWorkbook = Workbooks.Open(flowPlanMaster, false, false);
            destWorksheet = destWorkbook.Worksheets.Item["Charge Data"];
            logging.WriteLine(System.DateTime.Now + ":\t" + "Success in setting up src & dest Workbooks & Worksheets.");
            Console.WriteLine(System.DateTime.Now + ":\t" + "Success in setting up src & dest Workbooks & Worksheets.");

            //setup Excel enviroment, Stop calcuation before copy
            App.Application.DisplayAlerts = false;
            App.Application.Calculation = Excel.XlCalculation.xlCalculationManual;
            logging.WriteLine(System.DateTime.Now + ":\t" + "Successfully turned off application alerts and set workbook to manual calculation.");
            Console.WriteLine(System.DateTime.Now + ":\t" + "Successfully turned off application alerts and set workbook to manual calculation.");

            //Gather the size of the range to copy
            srcUsedRange = srcWorksheet.UsedRange.Rows.Count;
            destUsedRange = destWorksheet.UsedRange.Rows.Count;

            //clear the destination range so we do not get ghosting data
            Excel.Range r1 = destWorksheet.Cells[2, 1];
            Excel.Range r2 = destWorksheet.Cells[destUsedRange, 6];
            Excel.Range destRange = destWorksheet.Range[r1, r2];
            destRange.Value = "";
            destWorkbook.Worksheets.Item["SOS"].Cells[4, 9].value = System.DateTime.Today;
            TimeSpan start = new TimeSpan(1, 0, 0); // 1 o'clock
            TimeSpan end = new TimeSpan(12, 0, 0); //12 o'clock
            TimeSpan now = DateTime.Now.TimeOfDay;


            logging.WriteLine(System.DateTime.Now + ":\t" + "Successfully cleared destionation range.");
            Console.WriteLine(System.DateTime.Now + ":\t" + "Successfully cleared destionation range.");

            //set up enviroment for each FC
            string[] warehouses = { "AVP1", "CFC1", "DFW1", "EFC3", "WFC2" };
            string[] shifts = { "Days", "Nights", "" };
            string archivePath = @"\\cfc1afs01\Operations-Analytics\Projects\Flow Plan\BackUpArchive";
            int[] dprows = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            int[] cfc1dprows = { 31, 34, 37, 41, 71, 74, 77, 81, 31, 34, 37, 41 };
            int[] avp1dprows = { 25, 28, 31, 35, 59, 62, 65, 69, 25, 28, 31, 35 };
            int[] dfw1dprows = { 27, 30, 33, 37, 60, 63, 66, 70, 27, 30, 33, 37 };
            int[] wfc2dprows = { 30, 33, 36, 40, 67, 70, 73, 77, 30, 33, 36, 40 };
            int[] efc3dprows = { 25, 28, 31, 35, 59, 62, 65, 69, 25, 28, 31, 35 };
            string laborplanloc = @"\\chewy.local\bi\BI Community Content\Finance\Labor Models\";
            double[] laborplaninfo = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };


            foreach (string wh in warehouses )
            {
                a = wh;
                switch (wh)
                {
                    case "AVP1":
                        createdFileDestination = @"\\avp1afs01\Outbound\AVP Flow Plan\Blank Copy\";
                        archivePath = @"\\avp1afs01\Outbound\AVP Flow Plan\FlowPlanArchive\NotProcessed\";
                        avp1dprows.CopyTo(dprows, 0);
                        break;
                    case "CFC1":
                        createdFileDestination = @"\\cfc1afs01\Outbound\Master Wash\Blank Copy\";
                        archivePath = @"\\cfc1afs01\Outbound\Master Wash\FlowPlanArchive\NotProcessed\";
                        cfc1dprows.CopyTo(dprows, 0);
                        break;
                    case "DFW1":
                        createdFileDestination = @"\\dfw1afs01\Outbound\Blank Flow Plan\";
                        archivePath = @"\\dfw1afs01\Outbound\FlowPlanArchive\NotProcessed\";
                        dfw1dprows.CopyTo(dprows, 0);
                        break;
                    case "EFC3":
                        createdFileDestination = @"\\wh-pa-fs-01\OperationsDrive\Blank Flow Plan\";
                        archivePath = @"\\wh-pa-fs-01\OperationsDrive\FlowPlanArchive\NotProcessed\";
                        efc3dprows.CopyTo(dprows, 0);
                        break;
                    case "WFC2":
                        createdFileDestination = @"\\wfc2afs01\Outbound\Outbound Flow Planner\Blank Copy\";
                        archivePath = @"\\wfc2afs01\Outbound\Outbound Flow Planner\FlowPlanArchive\NotProcessed\";
                        wfc2dprows.CopyTo(dprows, 0);
                        break;
                    default:
                        Console.WriteLine("Warehouse not found please add to structure" + a);
                        logging.WriteLine("Warehouse not found please add to structure" + a);
                        break;
                 }

                


                if (runLaborPlan)
                {
                    runLaborPlanPopulate = obtainLaborPlanInfo(logging, laborplanloc, dprows, wh, Workbooks, laborplaninfo);
                }
                else
                {
                    runLaborPlanPopulate = false;
                }
               
                if (runArchive)
                {
                    CleanUpDirectories(createdFileDestination, archivePath, logging);
                }

                double[] avp1TPHHourlyGoal = {0.80,1.12,1.12,0.80,1.08,0.96,1.12,0.80,1.12,1.04};
                double[] OtherFCTPHHourlyGoal = { 0.80, 1.12, 0.80, 1.12, 1.08, 0.96, 1.12, 0.80, 1.12, 1.04 };
                foreach (string shift in shifts)
                {
                    //setup enviroment
                    App.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

                    //Define the destination workbook
                    destWorkbook = Workbooks.Open(flowPlanMaster, false, false);
                    destWorksheet = destWorkbook.Worksheets.Item["Charge Data"];
                    
                    //clear the destination range so we do not get ghosting data
                    r1 = destWorksheet.Cells[1, 1];
                    r2 = destWorksheet.Cells[destUsedRange, 6];
                    destRange = destWorksheet.Range[r1, r2];
                    destRange.Value = "";

                    //State which file we are creating
                    Console.WriteLine(System.DateTime.Now + ":\t" + "Starting " + a + "'s " + shift + " file preperation...");
                    logging.WriteLine(System.DateTime.Now + ":\t" + "Starting " + a + "'s " + shift + " file preperation...");

                    //Set SOS values
                    destWorkbook.Worksheets.Item["SOS"].Cells[3, 9].value = a;
                    destWorkbook.Worksheets.Item["SOS"].Cells[4, 9].value = System.DateTime.Now.ToString("yyyy-MM-dd");
                    destWorkbook.Worksheets.Item["SOS"].Cells[5, 9].value = shift;

                    for(int i=0; i < 10; i++)
                    {
                        if(a=="AVP1")
                        {
                            destWorkbook.Worksheets.Item["Hourly TPH"].Cells[25, i + 4].value = avp1TPHHourlyGoal[i];
                        }
                    }

                    if (runLaborPlanPopulate)
                    {
                        //set 21DP info
                        switch (shift)
                        {
                            case "Days":
                                destWorkbook.Worksheets.Item["SOS"].Cells[10, 3].value = laborplaninfo[0];//days hours
                                destWorkbook.Worksheets.Item["SOS"].Cells[11, 3].value = laborplaninfo[1];//days tph
                                destWorkbook.Worksheets.Item["SOS"].Cells[9, 3].value = laborplaninfo[2];//days shipped
                                destWorkbook.Worksheets.Item["SOS"].Cells[12, 3].value = laborplaninfo[3];//days ordered
                                destWorkbook.Worksheets.Item["SOS"].Cells[15, 3].value = laborplaninfo[4];//nights hours
                                destWorkbook.Worksheets.Item["SOS"].Cells[16, 3].value = laborplaninfo[5];//nights tph
                                destWorkbook.Worksheets.Item["SOS"].Cells[14, 3].value = laborplaninfo[6];//nights shipped
                                destWorkbook.Worksheets.Item["SOS"].Cells[17, 3].value = laborplaninfo[7];//nights ordered
                                break;
                            case "Nights":
                                destWorkbook.Worksheets.Item["SOS"].Cells[15, 3].value = laborplaninfo[8];//next days hours
                                destWorkbook.Worksheets.Item["SOS"].Cells[16, 3].value = laborplaninfo[9];//next days tph
                                destWorkbook.Worksheets.Item["SOS"].Cells[14, 3].value = laborplaninfo[10];//next days shipped
                                destWorkbook.Worksheets.Item["SOS"].Cells[17, 3].value = laborplaninfo[11];//next days ordered
                                destWorkbook.Worksheets.Item["SOS"].Cells[10, 3].value = laborplaninfo[4];//nights hours
                                destWorkbook.Worksheets.Item["SOS"].Cells[11, 3].value = laborplaninfo[5];//nights tph
                                destWorkbook.Worksheets.Item["SOS"].Cells[9, 3].value = laborplaninfo[6];//nights shipped
                                destWorkbook.Worksheets.Item["SOS"].Cells[12, 3].value = laborplaninfo[7];//nights ordered
                                break;
                            default:
                                break;
                        }
                    }
                   

                    //setup destination range varabiles
                    r1 = destWorksheet.Cells[1, 1];
                    r2 = destWorksheet.Cells[srcUsedRange, 6];
                    destRange = destWorksheet.Range[r1, r2];

                    //Copy source to it's destination
                    destRange.Value = srcWorksheet.UsedRange.Value;

                    Console.WriteLine(System.DateTime.Now + ":\t" + "Finishing " + a + "'s " + shift + " file preperation.");
                    logging.WriteLine(System.DateTime.Now + ":\t" + "Finishing " + a + "'s " + shift + " file preperation.");

                    filename = createdFileDestination + destWorkbook.Worksheets.Item["SOS"].cells(3, 9).value + " Flow Plan v2 " + shift + " " + System.DateTime.Now.ToString("yyyy-MM-dd");
                    App.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                    //Save as the workbook
                    var fi = new FileInfo(filename+".xlsm");
                    if (fi.Exists)
                    {
                        try
                        {
                            File.Delete(filename + ".xlsm");
                        }
                        catch (IOException)
                        {
                            Console.WriteLine(System.DateTime.Now + ":\t" + "Copy of " + a + " " + shift + " Flow Plan at " + createdFileDestination + " does not exist.");
                            logging.WriteLine(System.DateTime.Now + ":\t" + "Copy of " + a + " " + shift + " Flow Plan at " + createdFileDestination + " does not exist.");
                        }
                    }

                    try
                    {
                        destWorkbook.SaveAs(filename + ".xlsm");
                        destWorkbook.Close();
                        Console.WriteLine(System.DateTime.Now + ":\t" + "Saved copy of " + a + " " + shift + " Flow Plan at " + createdFileDestination);
                        logging.WriteLine(System.DateTime.Now + ":\t" + "Saved copy of " + a + " " + shift + " Flow Plan at " + createdFileDestination);
                    }
                    catch (Exception)
                    {
                        try
                        {
                            destWorkbook.SaveAs(filename + "(Empty).xlsm");
                            destWorkbook.Close();
                            Console.WriteLine(System.DateTime.Now + ":\t" + "Saved copy of " + a + " " + shift + " Flow Plan at " + createdFileDestination);
                            logging.WriteLine(System.DateTime.Now + ":\t" + "Saved copy of " + a + " " + shift + " Flow Plan at " + createdFileDestination);
                        }
                        catch
                        {
                            Console.WriteLine(System.DateTime.Now + ":\t" + "Unable to save any copy of " + a + " " + shift + " Flow Plan at " + createdFileDestination);
                            logging.WriteLine(System.DateTime.Now + ":\t" + "Unable to save any copy of " + a + " " + shift + " Flow Plan at " + createdFileDestination);
                        }
                        
                    }

                }
            }

           
            //null out the ranges
            r1 = null;
            r2 = null;
            destRange = null;
            
            //Restore calcuation
            App.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

            //close the worbooks saving destination
            srcWorkbook.Close(false);
            logging.WriteLine(System.DateTime.Now + ":\t" + "Successfully closed Raw Charge Pattern Data.");

            //Null out the remaining Excel variables
            destWorkbook = null;
            destWorksheet = null;
            srcWorkbook = null;
            srcWorksheet = null;

            //quit the excel application
            App.Quit();
            App = null;
            logging.WriteLine(System.DateTime.Now + ":\t" + "Successfully closed references to the Excel Application.");

            //clear com object references
            if (destWorkbook != null) 
            { Marshal.Marshal.ReleaseComObject(destWorkbook); }
            if ( destWorksheet != null)
            { Marshal.Marshal.ReleaseComObject(destWorksheet); }
            if (srcWorkbook != null)
            { Marshal.Marshal.ReleaseComObject(srcWorkbook); }
            if (srcWorksheet != null)
            { Marshal.Marshal.ReleaseComObject(srcWorksheet); }
            if ( Workbooks != null )
            { Marshal.Marshal.ReleaseComObject(Workbooks); }
            if ( App != null)
            { Marshal.Marshal.ReleaseComObject(App); }

            //Yell Success!
            Console.WriteLine("HORRAY Execution Completed!");
            logging.WriteLine(System.DateTime.Now + ":\t" + "Execution Completed");
            //close reference to streamwriter
            logging.Close();
            logging = null;
            
            
        }
        private static void CleanUpDirectories(string locToArchive, string archiveLocation, StreamWriter logging)
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
                    Console.WriteLine(System.DateTime.Now + ":\t" + "Directory confirmend at: " + dir.ToString());
                    logging.WriteLine(System.DateTime.Now + ":\t" + "Directory confirmend at: " + dir.ToString());
                }
                else
                {
                    Console.WriteLine(System.DateTime.Now + ":\t" + "Directory created at: " + dir.ToString());
                    logging.WriteLine(System.DateTime.Now + ":\t" + "Directory created at: " + dir.ToString());
                    System.IO.Directory.CreateDirectory(dir.ToString());
                }

                try
                {
                    Console.WriteLine(System.DateTime.Now + ":\t" + "File Coppied from: " + copy + " to Archive");
                    logging.WriteLine(System.DateTime.Now + ":\t" + "File Coppied from: " + copy + " to Archive");
                    System.IO.File.Copy(copy, archivePathFile, true);
                    System.IO.File.Delete(copy);
                }
                catch (System.IO.IOException)
                {
                    Console.WriteLine(System.DateTime.Now + ":\t" + "File currently in use: " + copy);
                    logging.WriteLine(System.DateTime.Now + ":\t" + "File currently in use: " + copy);
                }
            }
        }

        private static bool obtainLaborPlanInfo(StreamWriter logging, string laborplanloc, int[]dprows, string wh, Excel.Workbooks Workbooks, double[] laborplaninfo)
        {
            Console.WriteLine(System.DateTime.Now + ":\t" + "Begin Gathereing infromation form " + wh + " Labor Plan");
            logging.WriteLine(System.DateTime.Now + ":\t" + "Begin Gathereing infromation form " + wh + " Labor Plan");

            string[] laborplans = Directory.GetFiles(laborplanloc);
            string laborplanfileName = "";
            Excel.Workbook laborPlanModel = null;
            int col = 0;
            col = System.DateTime.Today.DayOfYear;
            int i = 0;
            int r = 0;
            string[] rowlabels = { "Planned Total Show Hours (Days)", "Planned Throughput (Days)", "Planned Units Shipped (Days)", "Planned Supply Chain Units Ordered (Days)", "Planned Total Show Hours (Nights)", "Planned Throughput (Nights)", "Planned Units Shipped (Nights)", "Planned Supply Chain Units Ordered (Nights)", "Planned Total Show Hours (Days)", "Planned Throughput (Days)", "Planned Units Shipped (Days)", "Planned Supply Chain Units Ordered (Days)" };


            foreach (string laborplan in laborplans)
            {
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
                        Console.WriteLine(System.DateTime.Now + ":\t" + "Unable to open flow plany at: " + laborplan);
                        logging.WriteLine(System.DateTime.Now + ":\t" + "Unable to open flow plany at: " + laborplan);
                        return false;
                    }
                    
                    Excel.Worksheet OBDP = laborPlanModel.Worksheets.Item["OB Daily Plan"];

                    //determine which rows to go look for
                    foreach (string label in rowlabels)
                    {
                        if(r>11)
                        {
                            Console.WriteLine(System.DateTime.Now + ":\t" + "Current Row Check Greater than number of verifications needed.. Shutting down Labor Plan Search");
                            logging.WriteLine(System.DateTime.Now + ":\t" + "Current Row Check Greater than number of verifications needed.. Shutting down Labor Plan Search");
                            return false;
                        }
                        string rcheck = OBDP.Cells[dprows[r], 2].value; //Set the row check = to what we think the row should be
                        if (rcheck == label)
                        {
                            Console.WriteLine(System.DateTime.Now + ":\t" + "Row location verified for " + label + " at row " + dprows[r]);
                            logging.WriteLine(System.DateTime.Now + ":\t" + "Row location verified for " + label + " at row " + dprows[r]);
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
                                    Console.WriteLine(System.DateTime.Now + ":\t" + "Row location for " + label + " now found at row " + dprows[r]);
                                    logging.WriteLine(System.DateTime.Now + ":\t" + "Row location for " + label + " now found at row " + dprows[r]);
                                }

                            }
                        }
                        r++;
                    }


                    foreach (int row in dprows)
                    {
                        if(i>11)
                        {
                            Console.WriteLine(System.DateTime.Now + ":\t" + "Current Row Check Greater than number of verifications needed.. Shutting down Labor Plan Search");
                            logging.WriteLine(System.DateTime.Now + ":\t" + "Current Row Check Greater than number of verifications needed.. Shutting down Labor Plan Search");
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
            return true; // we want to add the data to the flow planners
        }







    }

    
  
}
