/*
This DesignBuilder script automates upgrading EnergyPlus IDF files by replacing 
AirTerminal:SingleDuct:ConstantVolume:NoReheat (CAV:NoReheat) terminals in zones 
that also contain ZoneHVAC:TerminalUnit:VariableRefrigerantFlow (VRF) units. 

It substitutes them with AirTerminal:SingleDuct:Mixer objects to enable mixing
primary air (from central DOAS) with secondary VRF exhaust air at the zone inlet.
*/

using System.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{

    public class IdfFindAndReplace : ScriptBase, IScript
    {
        IdfReader Reader;
        const string ditributionUnitType = "ZoneHVAC:AirDistributionUnit";
        const string oldTerminalType = "AirTerminal:SingleDuct:ConstantVolume:NoReheat";
        const string newTerminalType = "AirTerminal:SingleDuct:Mixer";
        const string zoneUnitType = "ZoneHVAC:TerminalUnit:VariableRefrigerantFlow";

        public struct ZoneEquipmentSpecification
        {
            public string ZoneEquipmentListName;
            public string DistributionUnitName;
            public string TerminalType;
            public string TerminalName;
            public string ZoneUnitType;
            public string ZoneUnitName;
            public string ZoneInletNodeListName
            {
                get
                {
                    if (ZoneEquipmentListName != null)
                        return ZoneEquipmentListName.Replace(" Equipment", " Air Inlet Node List");
                    return null;
                }
            }
        }

        // Boilerplate for new Mixer terminal
        string terminalBoilerPlate = @"AirTerminal:SingleDuct:Mixer,
    {0},                   !- Name
    {1},                   !- ZoneHVAC Unit Object Type
    {2},                   !- ZoneHVAC Unit Object Name
    {3},                   !- Mixer Outlet Node Name
    {4},                   !- Mixer Primary Air Inlet Node Name
    {5},                   !- Mixer Secondary Air Inlet Node Name
    InletSide,             !- Mixer Connection Type
    {6},                   !- Design Specification Outdoor Air Object Name
    ;                      !- Per Person Ventilation Rate Mode";

        public override void BeforeEnergySimulation()
        {
            Reader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            ApplyMixedAirTerminals();
            Reader.Save();
        }

        // Main routine: replace terminals and update all linked objects
        public void ApplyMixedAirTerminals()
        {
            // Identify the zones containing the old terminal (CAV:NoReheat) and a VRF terminal unit
            List<ZoneEquipmentSpecification> specs = FindZoneEquipment(oldTerminalType, zoneUnitType);
            foreach (ZoneEquipmentSpecification spec in specs)
            {
                IdfObject oldTerminal = FindObject(spec.TerminalType, spec.TerminalName);
                IdfObject zoneUnit = FindObject(spec.ZoneUnitType, spec.ZoneUnitName);

                string exhaustNodeName = zoneUnit["Terminal Unit Air Inlet Node Name"].Value;

                // Build new Mixer object with required nodes
                string mixerText = String.Format(terminalBoilerPlate,
                    oldTerminal["Name"].Value,
                    spec.ZoneUnitType,
                    spec.ZoneUnitName,
                    oldTerminal["Air Outlet Node Name"].Value,
                    oldTerminal["Air Inlet Node Name"].Value,
                    exhaustNodeName,
                    oldTerminal["Design Specification Outdoor Air Object Name"].Value);

                IdfObject distributionUnit = FindObject(ditributionUnitType, spec.DistributionUnitName);
                distributionUnit["Air Terminal Object Type"].Value = newTerminalType;

                IdfObject coolingCoil = FindObject(zoneUnit["Cooling Coil Object Type"].Value, zoneUnit["Cooling Coil Object Name"].Value);
                coolingCoil["Coil Air Inlet Node"].Value = oldTerminal["Air Outlet Node Name"];

                IdfObject zoneNodeList = FindObject("NodeList", spec.ZoneInletNodeListName);
                RemoveTerminalOutletNode(zoneNodeList, oldTerminal["Air Outlet Node Name"].Value);

                // Insert Mixer and remove old terminal
                Reader.Load(mixerText);
                Reader.Remove(oldTerminal);

                zoneUnit["Terminal Unit Air Inlet Node Name"].Value = oldTerminal["Air Outlet Node Name"].Value;
            }
        }

        public IdfObject FindObject(string objectType, string objectName)
        {
            try
            {
                return Reader[objectType].First(c => c[0] == objectName);
            }
            catch (Exception e)
            {
                throw new MissingFieldException(String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        // Remove the terminal outlet node from the zone inlet
        public void RemoveTerminalOutletNode(IdfObject nodeList, string terminalOutletNodeName)
        {
            for (int i = 0; i < nodeList.Fields.Count; i++)
            {
                if (nodeList[i].Value == terminalOutletNodeName)
                {
                    nodeList.RemoveField(i);
                    break;
                }
            }
        }


        // Filter Zone equipments (terminal to be replaced + VRF terminal)
        public List<ZoneEquipmentSpecification> FindZoneEquipment(string terminalType, string zoneUnitType)
        {
            List<ZoneEquipmentSpecification> specifications = new List<ZoneEquipmentSpecification>();
            IEnumerable<IdfObject> allZoneEquipment = Reader["ZoneHVAC:EquipmentList"];
            foreach (IdfObject zoneEquipment in allZoneEquipment)
            {
                int i = 0;
                string zoneEquipmentListName = "";
                string terminalName = "";
                string zoneUnitName = "";
                string distributionUnitName = "";
                bool includesTerminal = false;
                bool includesZoneUnit = false;
                foreach (var field in zoneEquipment.Fields)
                {
                    if (field.Equals(ditributionUnitType))
                    {
                        distributionUnitName = zoneEquipment[i + 1].Value;
                        IdfObject distributionUnit = FindObject(ditributionUnitType, distributionUnitName);
                        string airTerminalType = distributionUnit["Air Terminal Object Type"].Value;
                        if (airTerminalType == terminalType)
                        {
                            terminalName = distributionUnit["Air Terminal Name"].Value;
                            includesTerminal = true;
                        }
                    }
                    if (field.Equals(zoneUnitType))
                    {
                        zoneUnitName = zoneEquipment[i + 1].Value;
                        includesZoneUnit = true;
                    }
                    i++;
                }
                if (includesTerminal && includesZoneUnit)
                {
                    ZoneEquipmentSpecification spec = new ZoneEquipmentSpecification
                    {
                        ZoneEquipmentListName = zoneEquipment[0].Value,
                        DistributionUnitName = distributionUnitName,
                        TerminalType = terminalType,
                        TerminalName = terminalName,
                        ZoneUnitType = zoneUnitType,
                        ZoneUnitName = zoneUnitName
                    };
                    specifications.Add(spec);
                }
            }
            return specifications;
        }
    }
}