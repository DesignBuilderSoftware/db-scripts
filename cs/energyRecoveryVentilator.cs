/*
Replace all zone exhaust fans with ERVs.

This DesignBuilder C# script edits the IDF to replace Fan:ZoneExhaust objects with ZoneHVAC:EnergyRecoveryVentilator (ERV) systems.

Purpose

1) Find all Fan:ZoneExhaust objects that are referenced inside ZoneHVAC:EquipmentList
2) Create a ZoneHVAC:EnergyRecoveryVentilator, its Controller, two Fan:SystemModel objects, and a HeatExchanger object per exhaust fan
3) Adjust the required NodeList
4) Replace the equipment reference in ZoneHVAC:EquipmentList from Fan:ZoneExhaust to ZoneHVAC:EnergyRecoveryVentilator
5) Remove the original Fan:ZoneExhaust objects from the IDF
6) Add a scheduled SetpointManager + NodeList and populating it with all heat exchanger supply outlet nodes

How to Use

Configuration
- Configure ERV parameters by editing the 'boilerplate' IDF text below:
   - ervBoilerplateIdf: ERV, controller, fans, heat exchanger performance/controls
   - spmBoilerplateIdf: setpoint manager / schedule / node list for HX outlet nodes
- ERV availability schedule and nominal air flow is read from Zone exhaust fan inputs in DesignBuilder.

Prerequisites (required placeholders)

- Ensure your model contains one or more Fan:ZoneExhaust objects assigned to zones
- The availability schedule and nominal air flow defined in the Fan:ZoneExhaust via user interface is applied in the ERV

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DB.Extensibility.Contracts;
using System;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        // ----------------------------
        // USER CONFIGURATION SECTION
        // ----------------------------
        // Boilerplate IDF used to create ERV + controller + fans + heat exchanger.
        string ervBoilerplateIdf = @"
ZoneHVAC:EnergyRecoveryVentilator,
    {0},                                      !- Name
    {1},                                      !- Availability Schedule Name
    {0} OA Heat Recovery,                     !- Heat Exchanger Name
    {2},                                      !- Supply Air Flow Rate
    {2},                                      !- Exhaust Air Flow Rate
    {0} Supply Fan,                           !- Supply Air Fan Name
    {0} Exhaust Fan,                          !- Exhaust Air Fan Name
    {0} OA Controller;                        !- Controller Name
	
ZoneHVAC:EnergyRecoveryVentilator:Controller,
    {0} OA Controller,                        !- Name
    ,                                         !- Temperature High Limit
    ,                                         !- Temperature Low Limit
    ,                                         !- Enthalpy High Limit
    ,                                         !- Dewpoint Temperature Limit
    ,                                         !- Electronic Enthalpy Limit Curve Name
    NoExhaustAirTemperatureLimit,             !- Exhaust Air Temperature Limit
    NoExhaustAirEnthalpyLimit,                !- Exhaust Air Enthalpy Limit
    {1},                                      !- Time of Day Economizer Flow Control Schedule Name
    No,                                       !- High Humidity Control Flag
    ,                                         !- Humidistat Control Zone Name
    ,                                         !- High Humidity Outdoor Air Flow Ratio
    No;                                       !- Control High Indoor Humidity Based on Outdoor Humidity Ratio
	
Fan:SystemModel,
    {0} Supply Fan,                           !- Name
    {1},                                      !- Availability Schedule Name
    {0} Heat Recovery Outlet Node,            !- Air Inlet Node Name
    {4},                                      !- Air Outlet Node Name
    {2},                                      !- Design Maximum Air Flow Rate
    Discrete,                                 !- Speed Control Method
    0.0,                                      !- Electric Power Minimum Flow Rate Fraction
    75.0,                                     !- Design Pressure Rise
    0.9,                                      !- Motor Efficiency
    1.0,                                      !- Motor In Air Stream Fraction
    AUTOSIZE,                                 !- Design Electric Power Consumption
    TotalEfficiencyAndPressure,               !- Design Power Sizing Method
    ,                                         !- Electric Power Per Unit Flow Rate
    ,                                         !- Electric Power Per Unit Flow Rate Per Unit Pressure
    0.50;                                     !- Fan Total Efficiency

Fan:SystemModel,
    {0} Exhaust Fan,                          !- Name
    {1},                                      !- Availability Schedule Name
    {0} Heat Recovery Secondary Outlet Node,  !- Air Inlet Node Name
    {0} Exhaust Fan Outlet Node,              !- Air Outlet Node Name
    {2},                                      !- Design Maximum Air Flow Rate
    Discrete,                                 !- Speed Control Method
    0.0,                                      !- Electric Power Minimum Flow Rate Fraction
    75.0,                                     !- Design Pressure Rise
    0.9,                                      !- Motor Efficiency
    1.0,                                      !- Motor In Air Stream Fraction
    AUTOSIZE,                                 !- Design Electric Power Consumption
    TotalEfficiencyAndPressure,               !- Design Power Sizing Method
    ,                                         !- Electric Power Per Unit Flow Rate
    ,                                         !- Electric Power Per Unit Flow Rate Per Unit Pressure
    0.50;                                     !- Fan Total Efficiency
	
HeatExchanger:AirToAir:SensibleAndLatent,
    {0} OA Heat Recovery,                     !- Name
    {1},                                      !- Availability Schedule Name
    {2},                                      !- Nominal Supply Air Flow Rate
    0.76,                                     !- Sensible Effectiveness at 100% Heating Air Flow
    0,                                        !- Latent Effectiveness at 100% Heating Air Flow
    0.81,                                     !- Sensible Effectiveness at 75% Heating Air Flow
    0,                                        !- Latent Effectiveness at 75% Heating Air Flow
    0.76,                                     !- Sensible Effectiveness at 100% Cooling Air Flow
    0,                                        !- Latent Effectiveness at 100% Cooling Air Flow
    0.81,                                     !- Sensible Effectiveness at 75% Cooling Air Flow
    0,                                        !- Latent Effectiveness at 75% Cooling Air Flow
    {5},                                      !- Supply Air Inlet Node Name
    {0} Heat Recovery Outlet Node,            !- Supply Air Outlet Node Name
    {3},                                      !- Exhaust Air Inlet Node Name
    {0} Heat Recovery Secondary Outlet Node,  !- Exhaust Air Outlet Node Name
    50.0,                                     !- Nominal Electric Power
    Yes,                                      !- Supply Air Outlet Temperature Control
    Plate,                                    !- Heat Exchanger Type
    MinimumExhaustTemperature,                !- Frost Control Type
    1.7;                                      !- Threshold Temperature";

        // ----------------------------
        // USER CONFIGURATION SECTION
        // ----------------------------
        // Boilerplate IDF used to add a setpoint manager/schedule and a NodeList to collect HX outlet nodes.
        string spmBoilerplateIdf = @"
SetpointManager:Scheduled,
    Heat Exchanger Supply Air Temp Manager,  !- Name
    Temperature,                             !- Control Variable
    Heat Exchanger Supply Air Temp Sch,      !- Schedule Name
    {0};                                     !- Setpoint Node or NodeList Name
	
NodeList,
    {0};                                     !- Name
	
Schedule:Compact,
    Heat Exchanger Supply Air Temp Sch,      !- Name
    Temperature,                             !- Schedule Type Limits Name
    Through: 12/31,                          !- Field 1
    For: AllDays,                            !- Field 2
    Until: 24:00,18;                         !- Field 3";

        private IdfObject FindObject(IdfReader idfReader, string objectType, string objectName)
        {
            return idfReader[objectType].First(o => o[0] == objectName);
        }

        // Scan ZoneHVAC:EquipmentList for entries of a given object type and return the referenced objects.
        private List<IdfObject> FindObjectsInZoneEquipment(IdfReader idfReader, string objectType)
        {
            List<IdfObject> objects = new List<IdfObject>();
            IEnumerable<IdfObject> allZoneEquipment = idfReader["ZoneHVAC:EquipmentList"];

            foreach (IdfObject zoneEquipment in allZoneEquipment)
            {
                int i = 0;

                foreach (var field in zoneEquipment.Fields)
                {
                    if (field.Equals(objectType))
                    {
                        string objectName = zoneEquipment[i + 1].Value;
                        IdfObject idfObject = FindObject(idfReader, objectType, objectName);
                        objects.Add(idfObject);
                    }
                    i++;
                }
            }
            return objects;
        }

        // Replace the matching (type,name) pairs in ZoneHVAC:EquipmentList.
        private void ReplaceObjectsInZoneEquipment(IdfReader idfReader, string oldObjectType, string oldObjectName, string newObjectType, string newObjectName)
        {
            IEnumerable<IdfObject> allZoneEquipment = idfReader["ZoneHVAC:EquipmentList"];

            bool objectFound = false;

            foreach (IdfObject zoneEquipment in allZoneEquipment)
            {
                if (!objectFound)
                {
                    for (int i = 0; i < (zoneEquipment.Count - 1); i++)
                    {
                        Field field = zoneEquipment[i];
                        Field nextField = zoneEquipment[i + 1];

                        if (field.Value == oldObjectType && nextField.Value == oldObjectName)
                        {
                            field.Value = newObjectType;
                            nextField.Value = newObjectName;
                            objectFound = true;
                            break;
                        }
                    }
                }
            }
        }

        // Collect node names from objects of a given type using the given field name.
        private List<string> FindNodes(IdfReader idfReader, string objectType, string fieldName)
        {
            List<string> nodes = new List<string>();
            IEnumerable<IdfObject> idfObjects = idfReader[objectType];

            foreach (IdfObject idfObject in idfObjects)
            {
                string nodeName = idfObject[fieldName].Value;

                if (nodeName.EndsWith("List"))
                {
                    IdfObject nodeList = idfReader["NodeList"].First(item => item[0] == nodeName);
                    nodeName = nodeList[1].Value;
                }
                nodes.Add(nodeName);
            }
            return nodes;
        }

        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // Target selection: only Fan:ZoneExhaust objects referenced in ZoneHVAC:EquipmentList are converted.
            IEnumerable<IdfObject> exhaustFans = FindObjectsInZoneEquipment(idfReader, "Fan:ZoneExhaust");

            string oldObjectType = "Fan:ZoneExhaust";
            string newObjectType = "ZoneHVAC:EnergyRecoveryVentilator";

            foreach (var exhaustFan in exhaustFans)
            {
                // Read key inputs from the exhaust fan placeholder.
                string exhaustFanName = exhaustFan[0].Value;
                string availabilityScheduleName = exhaustFan[1].Value;
                string nominalFlowRate = exhaustFan[4].Value;
                string zoneExhaustAirNode = exhaustFan[5].Value;
                string zoneName = exhaustFanName.Split(' ')[0];

                // Create new ERV and related nodes using a predictable naming convention.
                string ervName = zoneName + " ERV";
                string zoneInletNode = ervName + " Supply Fan Outlet Node";
                string ervOutdoorAirInletNode = ervName + " Supply ERV Inlet Node";

                // Add ERV + controller + fans + HX to the IDF.
                string ervIdfText = String.Format(
                    ervBoilerplateIdf,
                    ervName,
                    availabilityScheduleName,
                    nominalFlowRate,
                    zoneExhaustAirNode,
                    zoneInletNode,
                    ervOutdoorAirInletNode);

                idfReader.Load(ervIdfText);

                // Ensure the ERV OA inlet node is included in the global outdoor air node list.
                IdfObject outdoorAirNodeList = idfReader["OutdoorAir:NodeList"][0];
                outdoorAirNodeList.AddField(ervOutdoorAirInletNode);

                // Add the ERV supply outlet node to the zone air inlet node list.
                IdfObject zoneAirInletNodeList = FindObject(idfReader, "NodeList", zoneName + " Air Inlet Node List");
                zoneAirInletNodeList.AddField(zoneInletNode);

                ReplaceObjectsInZoneEquipment(idfReader, oldObjectType, exhaustFan[0].Value, newObjectType, ervName);

                // Remove the original exhaust fan object from the IDF.
                idfReader.Remove(exhaustFan);
            }

            // Add setpoint manager + schedule + node list for HX outlet nodes.
            string ervOutletNodeListName = "ERV HR Outlets";
            string spmObjects = String.Format(spmBoilerplateIdf, ervOutletNodeListName);
            idfReader.Load(spmObjects);
            List<string> heatExchangerSupplyOutletNodes = FindNodes(
                idfReader,
                "HeatExchanger:AirToAir:SensibleAndLatent",
                "Supply Air Outlet Node Name");

            IdfObject ervOutletNodeList = FindObject(idfReader, "NodeList", ervOutletNodeListName);
            ervOutletNodeList.AddFields(heatExchangerSupplyOutletNodes.ToArray());

            idfReader.Save();
        }
    }
}