/*
Add Dehumidifier:Desiccant:NoFans to an AirLoopHVAC outdoor air system (OA System).

Purpose
This DesignBuilder C# script inserts a desiccant dehumidifier component into an existing AirLoopHVAC outdoor air system.
The dehumidifier is inserted immediately upstream of an air-to-air heat exchanger in the AirLoopHVAC:OutdoorAirSystem:EquipmentList.

Main Steps

1) Locate the target AirLoopHVAC:OutdoorAirSystem:EquipmentList for a given air loop name
2) Identify the heat exchanger and the the equipment object immediately upstream
3) Insert Dehumidifier:Desiccant:NoFans into the equipment list before the heat exchanger
4) Rewire nodes so the upstream equipment outlet feeds the dehumidifier, and the dehumidifier feeds the heat exchanger
5) Create and load required supporting objects:
  - Schedule:Compact (availability)
  - Fan:VariableVolume (regeneration fan)
  - Coil:Heating:Electric (regeneration heater)
  - SetpointManager:MultiZone:Humidity:Maximum (humidity setpoint control)

How to Use

Configuration

Defined in the AddDesiccantDehumidifierToOaSystem():
- airLoopName: Must match the model's IDF naming.
- heatExchangerEquipmentFieldIndex: The field index where the heat exchanger object type appears in the equipment list

Prerequisites (required placeholders)

This script expects an Air Loop of type AirLoopHVAC:OutdoorAirSystem:EquipmentList to be in place in the model.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Linq;
using System.Text;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class AddDesiccantDehumidifierToOaSystemScript : ScriptBase, IScript
    {
        // IDF text templates for objects created by this script.
        private readonly string dehumidifierIdfTemplate = @"
Dehumidifier:Desiccant:NoFans,
    {0},                     !- Name
    {1},                     !- Availability Schedule Name
    {2},                     !- Process Air Inlet Node Name
    {3},                     !- Process Air Outlet Node Name
    Desiccant Heating Coil Air Outlet Node,     !- Regeneration Air Inlet Node Name
    {5},                     !- Regeneration Fan Inlet Node Name
    SystemNodeMaximumHumidityRatioSetpoint,  !- Control Type alt LeavingMaximumHumidityRatioSetpoint
    0.007,                   !- Leaving Maximum Humidity Ratio Setpoint kgWater/kgDryAir
    1.5,                     !- Nominal Process Air Flow Rate m3/s
    2.5,                     !- Nominal Process Air Velocity m/s
    50,                      !- Rotor Power W
    {4},                     !- Regeneration Coil Object Type
    {6},                     !- Regeneration Coil Name
    Fan:VariableVolume,      !- Regeneration Fan Object Type
    {7},                     !- Regeneration Fan Name
    DEFAULT;                 !- Performance Model Type

OutdoorAir:Node, Outside Air Inlet Node 2;";

        private readonly string regenFanIdfTemplate = @"
Fan:VariableVolume,
    {0},                     !- Name
    {1},                     !- Availability Schedule Name
    0.7,                     !- Fan Total Efficiency
    600.0,                   !- Pressure Rise Pa
    Autosize,                !- Maximum Flow Rate m3/s
    FixedFlowRate,           !- Fan Power Minimum Flow Rate Input Method
    ,                        !- Fan Power Minimum Flow Fraction
    0.0,                     !- Fan Power Minimum Air Flow Rate m3/s
    0.9,                     !- Motor Efficiency
    1.0,                     !- Motor In Airstream Fraction
    0,                       !- Fan Power Coefficient 1
    1,                       !- Fan Power Coefficient 2
    0,                       !- Fan Power Coefficient 3
    0,                       !- Fan Power Coefficient 4
    0,                       !- Fan Power Coefficient 5
    {2},                     !- Air Inlet Node Name
    Regen Fan Outlet Node;   !- Air Outlet Node Name";

        private readonly string regenCoilIdfTemplate = @"
Coil:Heating:Electric,
    {0},                     !- Name
    {1},                     !- Availability Schedule Name
    1.0,                     !- Efficiency
    autosize,                !- Nominal Capacity W
    Regen Fan Outlet Node,   !- Air Inlet Node Name
    Desiccant Heating Coil Air Outlet Node;     !- Air Outlet Node Name";

        private readonly string availabilityScheduleIdfTemplate = @"
Schedule:Compact, {0},       ! Name
   Any Number,               ! Type
   Through: 12/31,           ! Type
   For: AllDays,             ! All days in year
   Until: 24:00,             ! All hours in day
   1;";

        private readonly string humiditySpmIdfTemplate = @"  
SetpointManager:MultiZone:Humidity:Maximum,
   Humidity Setpoint Manager,                           ! - Component name
   {0},                                                 ! - HVAC air loop name
   .005,                                                ! - Minimum setpoint humidity ratio (kg/kg)
   .012,                                                ! - Maximum setpoint humidity ratio (kg/kg)
   {1};                                                 ! - Setpoint node list";

        // Helper to find an IDF object by type and name.
        private IdfObject FindObject(IdfReader reader, string objectType, string objectName)
        {
            try
            {
                return reader[objectType].First(c => c[0] == objectName);
            }
            catch (Exception)
            {
                throw new Exception(string.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        // Inserts a desiccant dehumidifier into the OA equipment list upstream of the heat exchanger
        private void AddDesiccantDehumidifierToOaSystem(IdfReader reader, string airLoopName, int heatExchangerEquipmentFieldIndex)
        {
            // The air loop outdoor air system equipment list will include the desiccant dehumidifier
            string oaEquipmentListName = airLoopName + " AHU Outdoor air Equipment List";
            IdfObject equipmentList = FindObject(reader, "AirLoopHVAC:OutdoorAirSystem:EquipmentList", oaEquipmentListName);

            // Find the air loop heat exchanger
            string heatExchangerObjectType = equipmentList[heatExchangerEquipmentFieldIndex].Value;
            string heatExchangerName = equipmentList[heatExchangerEquipmentFieldIndex + 1].Value;
            IdfObject heatExchanger = FindObject(reader, heatExchangerObjectType, heatExchangerName);

            // The desiccant unit is inserted immediately before the heat exchanger
            int upstreamEquipmentFieldIndex = heatExchangerEquipmentFieldIndex - 2;
            string upstreamEquipmentObjectType = equipmentList[upstreamEquipmentFieldIndex].Value;
            string upstreamEquipmentName = equipmentList[upstreamEquipmentFieldIndex + 1].Value;
            IdfObject upstreamEquipment = FindObject(reader, upstreamEquipmentObjectType, upstreamEquipmentName);

            // Specify process inlet and outlet nodes
            string processAirInletNode = upstreamEquipment["Air Outlet Node Name"].Value;
            string processAirOutletNode = airLoopName + " Desiccant Process Outlet Node";

            // The heat exchanger supply inlet node will be connected to the desiccant dehumidifier 
            heatExchanger["Supply Air Inlet Node Name"].Value = processAirOutletNode;

            // Insert the dehumidifier in the OA equipment list (as a type/name pair).
            string desiccantDehumidifierObjectType = "Dehumidifier:Desiccant:NoFans";
            string desiccantDehumidifierName = airLoopName + " Desiccant Unit";
            equipmentList.InsertFields(
                upstreamEquipmentFieldIndex,
                new string[] { desiccantDehumidifierObjectType, desiccantDehumidifierName });

            // Define the regeneration inlet node, this can be either the air loop relief or outdoor air node
            string regenAirInletNode = heatExchanger["Exhaust Air Outlet Node Name"].Value;

            // Load all desiccant dehumidifer related objects
            string scheduleName = airLoopName + " Desiccant unit schedule";
            string availabilitySchedule = string.Format(availabilityScheduleIdfTemplate, scheduleName);

            string regenFanName = airLoopName + " Desiccant Fan";
            string regenFan = string.Format(regenFanIdfTemplate, regenFanName, scheduleName, regenAirInletNode);

            string regenCoilName = airLoopName + " Desiccant Heating Coil";
            string regenCoil = string.Format(regenCoilIdfTemplate, regenCoilName, scheduleName);
            string regenCoilObjectType = "Coil:Heating:Electric";

            string humiditySpm = string.Format(humiditySpmIdfTemplate, airLoopName, processAirOutletNode);

            // Build the dehumidifier object.
            string dehumidifier = string.Format(
                dehumidifierIdfTemplate,
                desiccantDehumidifierName,
                scheduleName,
                processAirInletNode,
                processAirOutletNode,
                regenCoilObjectType,
                regenAirInletNode,
                regenCoilName,
                regenFanName);

            reader.Load(dehumidifier);
            reader.Load(availabilitySchedule);
            reader.Load(regenFan);
            reader.Load(regenCoil);
            reader.Load(humiditySpm);
        }

        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // ----------------------------
            // USER CONFIGURATION SECTION
            // ----------------------------
            // "Air loop" must match the air loop name used to form the OA equipment list name.
            // Field index (e.g., 5) is where the heat exchanger object type appears in the equipment list.
            AddDesiccantDehumidifierToOaSystem(idfReader, "Air loop", 5);

            idfReader.Save();
        }
    }
}