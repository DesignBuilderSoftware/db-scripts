/*
Replace all Fan:ZoneExhaust with ZoneHVAC:EnergyRecoveryVentilators (ERV).

ERV availability schedule and nominal air flow is read from Zone exhaust fan inputs in DesignBuilder.
Remaining attributes can be set in 'boilerplate' IDF text below.
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
        string ervBoilerPlate = @"
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

        string spmBoilerPlate = @"
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

        private List<string> FindNodes(IdfReader idfReader, string objectName, string fieldName)
        {
            List<string> nodes = new List<string>();
            IEnumerable<IdfObject> idfObjects = idfReader[objectName];

            foreach (IdfObject idfObject in idfObjects)
            {
                string nodeName = idfObject[fieldName];

                if (nodeName.EndsWith("List"))
                {
                    IdfObject nodeList = idfReader["NodeList"].First(item => item[0] == nodeName);
                    nodeName = nodeList[1];
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

            IEnumerable<IdfObject> exhaustFans = FindObjectsInZoneEquipment(idfReader, "Fan:ZoneExhaust");

            string oldObjectType = "Fan:ZoneExhaust";
            string newObjectType = "ZoneHVAC:EnergyRecoveryVentilator";

            foreach (var exhaustFan in exhaustFans)
            {
                string name = exhaustFan[0].Value;
                string schedule = exhaustFan[1].Value;
                string flowRate = exhaustFan[4].Value;
                string zoneExhaustNode = exhaustFan[5].Value;
                string zoneName = name.Split(' ')[0];
                string ervName = name.Split(' ')[0] + " ERV";
                string zoneInletNode = ervName + " Supply Fan Outlet Node";
                string ervOANode = ervName + " Supply ERV Inlet Node";

                string newComponents = String.Format(ervBoilerPlate, ervName, schedule, flowRate, zoneExhaustNode, zoneInletNode, ervOANode);
                idfReader.Load(newComponents);

                IdfObject oaNodes = idfReader["OutdoorAir:NodeList"][0];
                oaNodes.AddField(ervOANode);

                IdfObject nodeList = FindObject(idfReader, "NodeList", zoneName + " Air Inlet Node List");
                nodeList.AddField(zoneInletNode);

                ReplaceObjectsInZoneEquipment(idfReader, oldObjectType, exhaustFan[0], newObjectType, ervName);

                idfReader.Remove(exhaustFan);
            }

            string nodeListName = "ERV HR Outlets";
            string spmObjects = String.Format(spmBoilerPlate, nodeListName);
            idfReader.Load(spmObjects);
            List<string> outletNodes = FindNodes(idfReader, "HeatExchanger:AirToAir:SensibleAndLatent", "Supply Air Outlet Node Name");

            IdfObject ervNodeList = FindObject(idfReader, "NodeList", nodeListName);
            ervNodeList.AddFields(outletNodes.ToArray());

            idfReader.Save();
        }
    }
}