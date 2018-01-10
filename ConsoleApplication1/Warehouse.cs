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
        private string _distroList;
        private bool _preShiftFlag;
        private bool _laborPlanInfoPop;
        private bool _preShiftInfoPop;
        private bool _IBFlowPlan;
        

        private double _handOffPercentage;
        private double _VCPUwageRate;
        private int _handOffDeadMan;

        private int[] _laborPlanInforRows;
        private double[] _tphdistrobution;
        private double[] _daysStaffingRates;
        private double[] _nightsStaffingRates;
        private double[] _daysChargePattern;
        private double[] _nightsChargePattern;


        public Warehouse(string name, string blankCopyLoc,
            string archiveLoc, int[] laborPlanInfoRows, double[] daysRates,
            double[] nightsRates, double[] tphDistro, string location = "Unknown",
            string distrolist = "camos@chewy.com", bool preShiftFlag = true,
            bool laborPlanInfoPop = true, bool preShiftInfoPop = true,
            double handOffPercentage = .46, int handOffDeadman = 24000, bool IBFlowPlan = true)
        {
            _name = name;
            _location = location;
            _OBblankCopyLoc = blankCopyLoc;
            _OBarchiveLocation = archiveLoc;
            _distroList = distrolist;
            _laborPlanInforRows = laborPlanInfoRows;
            _tphdistrobution = tphDistro;
            _daysStaffingRates = daysRates;
            _nightsStaffingRates = nightsRates;
            _preShiftFlag = preShiftFlag;
            _preShiftInfoPop = preShiftInfoPop;
            _laborPlanInfoPop = laborPlanInfoPop;
            _handOffPercentage = HandoffPercent;
            _handOffDeadMan = handOffDeadman;
            _IBFlowPlan = IBFlowPlan;
        }
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
        public string DistroList
        {
            get { return _distroList; }
            set { _distroList = value; }
        }
        public int[] laborPlanInforRows
        {
            get { return _laborPlanInforRows; }
            set { _laborPlanInforRows = value; }
        }
        public double[] tphDistro
        {
            get { return _tphdistrobution; }
            set { _tphdistrobution = value; }
        }
        public double[] daysRate
        {
            get { return _daysStaffingRates; }
            set { _daysStaffingRates = value; }
        }
        public double[] nightsRate
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
        public bool PreShiftFlag
        {
            get { return _preShiftFlag; }
            set { _preShiftFlag = value; }
        }
        public int DeadMan
        {
            get { return _handOffDeadMan; }
            set { _handOffDeadMan = value; }
        }
        public bool laborInfoPop
        {
            get { return _laborPlanInfoPop; }
            set { _laborPlanInfoPop = value; }
        }
        public bool preshiftInfoPop
        {
            get { return _preShiftInfoPop; }
            set { _preShiftInfoPop = value; }
        }
        public double[] daysChargePattern
        {
            get { return _daysChargePattern; }
            set { _daysChargePattern = value; }
        }
        public double[] nightsChargePattern
        {
            get { return _nightsChargePattern; }
            set { _nightsChargePattern = value; }
        }
        public void buildDaysChargePattern(double[] srcData)
        {

        }
        public bool IBFlowPlan
        {
            get { return _IBFlowPlan; }
            set { _IBFlowPlan = value; }
        }
    }
}
