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
        private string _blankCopyLoc;
        private string _archiveLocation;
        private string _distroList;

        private double _handOffPercentage;
        private double _VCPUwageRate;

        private int[] _laborPlanInforRows;
        private double[] _tphdistrobution;
        private double[] _daysStaffingRates;
        private double[] _nightsStaffingRates;
       

        public Warehouse(string name, string location, string blankCopyLoc, string archiveLoc, string distrolist, int[] laborPlanInfoRows, double[] tphDistro, double[] daysRates, double[] nightsRates)
        {
            _name = name;
            _location = location;
            _blankCopyLoc = blankCopyLoc;
            _archiveLocation = archiveLoc;
            _distroList = distrolist;
            _laborPlanInforRows = laborPlanInfoRows;
            _tphdistrobution = tphDistro;
            _daysStaffingRates = daysRates;
            _nightsStaffingRates = nightsRates;
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
        public string blankCopyLoc
        {
            get { return _blankCopyLoc; }
            set { _blankCopyLoc = value; }
        }
        public string archiveLoc
        {
            get { return _archiveLocation; }
            set { _archiveLocation = value; }
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
        //public static Warehouse CreatewarehouseFromXmlString(string xmlWarehouseData)
        //{
        //try
        //{
        //    XmlDocument warehouseData = new XmlDocument();

        //    warehouseData.LoadXml(xmlWarehouseData);

        //    int currentHitPoints = Convert.ToInt32(warehouseData.SelectSingleNode("/Player/Stats/CurrentHitPoints").InnerText);
        //    int maximumHitPoints = Convert.ToInt32(warehouseData.SelectSingleNode("/Player/Stats/MaximumHitPoints").InnerText);
        //    int gold = Convert.ToInt32(warehouseData.SelectSingleNode("/Player/Stats/Gold").InnerText);
        //    int experiencePoints = Convert.ToInt32(warehouseData.SelectSingleNode("/Player/Stats/ExperiencePoints").InnerText);

        //    Warehouse warehouses = new Warehouse(currentHitPoints, maximumHitPoints, gold, experiencePoints);

        //    int currentLocationID = Convert.ToInt32(warehouseData.SelectSingleNode("/Player/Stats/CurrentLocation").InnerText);

        //    //if (warehouseData.SelectSingleNode("/Player/Stats/CurrentWeapon") != null)
        //    //{
        //    //    int currentWeaponID = Convert.ToInt32(warehouseData.SelectSingleNode("/Player/Stats/CurrentWeapon").InnerText);
        //    //    player.CurrentWeapon = (Weapon)World.ItemByID(currentWeaponID);
        //    //}

        //    //foreach (XmlNode node in warehouseData.SelectNodes("/Player/InventoryItems/InventoryItem"))
        //    //{
        //    //    int id = Convert.ToInt32(node.Attributes["ID"].Value);
        //    //    int quantity = Convert.ToInt32(node.Attributes["Quantity"].Value);

        //    //    for (int i = 0; i < quantity; i++)
        //    //    {
        //    //        player.AddItemToInventory(World.ItemByID(id));
        //    //    }
        //    //}

        //    foreach (XmlNode node in warehouseData.SelectNodes("/Player/PlayerQuests/PlayerQuest"))
        //    {
        //        int id = Convert.ToInt32(node.Attributes["ID"].Value);
        //        bool isCompleted = Convert.ToBoolean(node.Attributes["IsCompleted"].Value);

        //        PlayerQuest playerQuest = new PlayerQuest(World.QuestByID(id));
        //        playerQuest.IsCompleted = isCompleted;

        //        warehouses.Quests.Add(playerQuest);
        //    }

        //    return warehouses;
        //}
        //catch
        //{
        //    // If there was an error with the XML data, return a default player object
        //    return Player.CreateDefaultPlayer();

        //}
        //}

        public string ToXmlString()
        {
            XmlDocument warehouseData = new XmlDocument();

            // Create the "Stats" child node to hold the other player statistics nodes
            XmlNode warehouse = warehouseData.CreateElement(_name);
            warehouseData.AppendChild(warehouse);

            AddXmlAttributeToNode(warehouseData, warehouse, "ID", 1);

            // Create the child nodes for the "warehouse" node
            CreateNewChildXmlNode(warehouseData, warehouse, "Location", _location.ToString());
            CreateNewChildXmlNode(warehouseData, warehouse, "BlankLoc", _blankCopyLoc.ToString());
            CreateNewChildXmlNode(warehouseData, warehouse, "ArchiveLoc", _archiveLocation.ToString());
            CreateNewChildXmlNode(warehouseData, warehouse, "DistroListTarget", _distroList.ToString());
            CreateNewChildXmlNode(warehouseData, warehouse, "LaborPlanRows", string.Join(",",_laborPlanInforRows));
            CreateNewChildXmlNode(warehouseData, warehouse, "TPHDistro", string.Join(",", _tphdistrobution));
            CreateNewChildXmlNode(warehouseData, warehouse, "DaysRates", string.Join(",", _daysStaffingRates));
            CreateNewChildXmlNode(warehouseData, warehouse, "NightsRates", string.Join(",", _nightsStaffingRates));

            return warehouseData.InnerXml; // The XML document, as a string, so we can save the data to disk
        }

        private void CreateNewChildXmlNode(XmlDocument document, XmlNode parentNode, string elementName, object value)
        {
            XmlNode node = document.CreateElement(elementName);
            node.AppendChild(document.CreateTextNode(value.ToString()));
            parentNode.AppendChild(node);
        }

        private void AddXmlAttributeToNode(XmlDocument document, XmlNode node, string attributeName, object value)
        {
            XmlAttribute attribute = document.CreateAttribute(attributeName);
            attribute.Value = value.ToString();
            node.Attributes.Append(attribute);
        }
    }
}
