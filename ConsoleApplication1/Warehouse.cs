using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace FlowPlanConstruction
{
    public class Warehouse
    {
        private string _name;
        private string _location;
        private string _OBblankCopyLoc;
        private string _OBarchiveLocation;
        private string _IBblankCopyLoc;
        private string _IBarchiveLocation;
        private string _CycleTimeToolLoc;
        private string _CycleTimeToolarchiveLoc;
        private string _distroList;
        private string _LSWCLoc;
        private bool _preShiftFlag;
        private bool _laborPlanInfoPop;
        private bool _preShiftInfoPop;
        private bool _IBFlowPlan;
        private bool _CycleTimeTool;
        private bool _copyTestPlan;


        private double _handOffPercentage;
        private double _VCPUwageRate;
        private int _handOffDeadMan;
        private int _daysrunhour;
        private int _nightsrunhour;
        private double _timeoffset;

        private int[] _laborPlanInforRows;
        private double[] _daystphdistrobution;
        private double[] _nightstphdistrobution;
        private double[] _daysStaffingRates;
        private double[] _nightsStaffingRates;
        //private double[] _daysChargePattern;
        //private double[] _nightsChargePattern;
        private double[] _mshiftsplit;
        private int[] _cutTimes;
        private double[] _capacityCalcRates;
        private double[] _daysBreakhours;
        private double[] _nightsBreakhours;

        
        //public Warehouse(string name, string blankCopyLoc,
        //    string archiveLoc, int[] laborPlanInfoRows, double[] daysRates,
        //    double[] nightsRates, double[] tphDistro, string location = "Unknown",
        //    string distrolist = "camos@chewy.com", bool preShiftFlag = true,
        //    bool laborPlanInfoPop = true, bool preShiftInfoPop = true,
        //    double handOffPercentage = .46, int handOffDeadman = 24000, bool IBFlowPlan = true, double Timeoffset = 0)
        //{
        //    _name = name;
        //    _location = location;
        //    _OBblankCopyLoc = blankCopyLoc;
        //    _OBarchiveLocation = archiveLoc;
        //    _distroList = distrolist;
        //    _laborPlanInforRows = laborPlanInfoRows;
        //    _daystphdistrobution = tphDistro;
        //    _daysStaffingRates = daysRates;
        //    _nightsStaffingRates = nightsRates;
        //    _preShiftFlag = preShiftFlag;
        //    _preShiftInfoPop = preShiftInfoPop;
        //    _laborPlanInfoPop = laborPlanInfoPop;
        //    _handOffPercentage = HandoffPercent;
        //    _handOffDeadMan = handOffDeadman;
        //    _IBFlowPlan = IBFlowPlan;
        //}
        public string Name
        {
            get { return _name; }
            set {  _name = value; }
        }
        public string Location
        {
            get { return _location; }
            set { _location = value; }
        }
        public int DaysRunTime
        {
            get { return _daysrunhour; }
            set { _daysrunhour = value; }
        }
        public int NightsRunTime
        {
            get { return _nightsrunhour; }
            set { _nightsrunhour = value; }
        }
        public bool IBFlowPlan
        {
            get { return _IBFlowPlan; }
            set { _IBFlowPlan = value; }
        }
        public bool CycleTimeTool
        {
            get { return _CycleTimeTool; }
            set { _CycleTimeTool = value; }
        }
        public bool LaborInfoPop
        {
            get { return _laborPlanInfoPop; }
            set { _laborPlanInfoPop = value; }
        }
        public bool PreshiftInfoPop
        {
            get { return _preShiftInfoPop; }
            set { _preShiftInfoPop = value; }
        }
        public bool PreShiftFlag
        {
            get { return _preShiftFlag; }
            set { _preShiftFlag = value; }
        }
        public bool CopyTestFile
        {
            get { return _copyTestPlan; }
            set { _copyTestPlan = value; }
        }
        public string LSWCLoc
        {
            get { return _LSWCLoc; }
            set { _LSWCLoc = value; }
        }
        public string OBblankCopyLoc
        {
            get { return _OBblankCopyLoc; }
            set { _OBblankCopyLoc = value; }
        }
        public string OBarchiveLoc
        {
            get { return _OBarchiveLocation; }
            set { _OBarchiveLocation = value; }
        }
        public string IBblankCopyLoc
        {
            get { return _IBblankCopyLoc; }
            set { _IBblankCopyLoc = value; }
        }
        public string IBarchiveLoc
        {
            get { return _IBarchiveLocation; }
            set { _IBarchiveLocation = value; }
        }
        public string CycleTimeToolBlankLoc
        {
            get { return _CycleTimeToolLoc; }
            set { _CycleTimeToolLoc = value; }
        }
        public string CycleTimeToolArcLoc
        {
            get { return _CycleTimeToolarchiveLoc; }
            set { _CycleTimeToolarchiveLoc = value; }
        }
        public string OBDistroList
        {
            get { return _distroList; }
            set { _distroList = value; }
        }
        public int[] LaborPlanInforRows
        {
            get { return _laborPlanInforRows; }
            set { _laborPlanInforRows = value; }
        }
        public double[] DaystphDistro
        {
            get { return _daystphdistrobution; }
            set { _daystphdistrobution = value; }
        }
        public double[] NightstphDistro
        {
            get { return _nightstphdistrobution; }
            set { _nightstphdistrobution = value; }
        }
        public double[] DaysBreaks
        {
            get { return _daysBreakhours; }
            set { _daysBreakhours = value; }
        }
        public double[] NightsBreaks
        {
            get { return _nightsBreakhours; }
            set { _nightsBreakhours = value; }
        }
        public double[] DaysRate
        {
            get { return _daysStaffingRates; }
            set { _daysStaffingRates = value; }
        }
        public double[] NightsRate
        {
            get { return _nightsStaffingRates; }
            set { _nightsStaffingRates = value; }
        }
        public double HandoffPercent
        {
            get { return _handOffPercentage;  }
            set { _handOffPercentage = value; }
        }
        public double VCPUWageRate
        {
            get { return _VCPUwageRate; }
            set { _VCPUwageRate = value; }
        }
        public int DeadMan
        {
            get { return _handOffDeadMan; }
            set { _handOffDeadMan = value; }
        }
        public double [] Mshiftsplit
        {
            get { return _mshiftsplit; }
            set { _mshiftsplit = value; }
        }
        public double Timeoffset
        {
            get { return _timeoffset; }
            set { _timeoffset = value; }
        }
        public int[] CutTimes
        {
            get { return _cutTimes; }
            set { _cutTimes = value; }
        }
        public double[] CapacityCalcRates
        {
            get { return _capacityCalcRates; }
            set { _capacityCalcRates = value; }
        }


        //public void BuildDaysChargePattern(double[] srcData)
        //{

        //}
        //public double[] DaysChargePattern
        //{
        //    get { return _daysChargePattern; }
        //    set { _daysChargePattern = value; }
        //}
        //public double[] NightsChargePattern
        //{
        //    get { return _nightsChargePattern; }
        //    set { _nightsChargePattern = value; }
        //}





    }
}
