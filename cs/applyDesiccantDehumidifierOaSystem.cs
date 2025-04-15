/*
Add Dehumidifier:Desiccant:NoFans object into an air handling unit.

The component is right before the air loop heat exchanger. 
Note that it's needed to specify the index of the heat exchanger position in the OA equipment list.

*/
using System.Runtime;
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
        string dehumidifierBoilerplate = @"  Dehumidifier:Desiccant:NoFans,
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

        string dehumidifierFanBoilerplate = @"  Fan:VariableVolume,
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

        string dehumidifierCoilBoilerplate = @"  Coil:Heating:Electric,
    {0},                     !- Name
    {1},                     !- Availability Schedule Name
    1.0,                     !- Efficiency
    autosize,                !- Nominal Capacity W
    Regen Fan Outlet Node,   !- Air Inlet Node Name
    Desiccant Heating Coil Air Outlet Node;     !- Air Outlet Node Name";

        string dehumidifierScheduleBoilerplate = @"  Schedule:Compact, {0},       ! Name
   Any Number,               ! Type
   Through: 12/31,           ! Type
   For: AllDays,             ! All days in year
   Until: 24:00,             ! All hours in day
   1;";

        string dehumidifierSpmBoilerplate = @"  SetpointManager:MultiZone:Humidity:Maximum,
   Humidity Setpoint Manager,                           ! - Component name
   {0},                                                 ! - HVAC air loop name
   .005,                                                ! - Minimum setpoint humidity ratio (kg/kg)
   .012,                                                ! - Maximum setpoint humidity ratio (kg/kg)
   {1};                                                 ! - Setpoint node list";


        public IdfObject FindObject(IdfReader reader, string objectType, string objectName)
        {
            try
            {
                return reader[objectType].First(c => c[0] == objectName);
            }
            catch (Exception e)
            {
                throw new Exception(String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        public void AddDesiccantDehumifierToOaSystem(IdfReader reader, string airLoopName, int hxEquipmentIndex)
        {
            // The air loop outdoor air system equipment list will include the desiccant dehumidifier
            string oaEquipmentListName = airLoopName + " AHU Outdoor air Equipment List";
            IdfObject equipmentList = FindObject(reader, "AirLoopHVAC:OutdoorAirSystem:EquipmentList", oaEquipmentListName);

            // Find the air loop heat exchanger
            string hxType = equipmentList[hxEquipmentIndex].Value;
            string hxName = equipmentList[hxEquipmentIndex + 1].Value;
            IdfObject hxObject = FindObject(reader, hxType, hxName);

            // The desiccant dehumidifier will be placed right before the heat exchanger
            int originalObjectIndex = hxEquipmentIndex - 2;
            string originalObjectType = equipmentList[originalObjectIndex].Value;
            string originalObjectname = equipmentList[originalObjectIndex + 1].Value;
            IdfObject originalObject = FindObject(reader, originalObjectType, originalObjectname);

            // Specify process inlet and outlet nodes
            string processInletNode = originalObject["Air Outlet Node Name"].Value;
            string processOutletNode = airLoopName + " Desiccant Process Outlet Node";

            // The heat exchanger supply inlet node will be connected to the desiccant dehumidifier 
            hxObject["Supply Air Inlet Node Name"].Value = processOutletNode;

            // Add the desiccant dehumidifier to the air loop outdoor air equipment list
            string desiccantDehumidifierComponent = "Dehumidifier:Desiccant:NoFans";
            string desiccantDehumidifierName = airLoopName + " Desiccant Unit";
            string[] newEquipmentFields = new string[] { desiccantDehumidifierComponent, desiccantDehumidifierName };
            equipmentList.InsertFields(originalObjectIndex, newEquipmentFields);

            // Define the regeneration inlet node, this can be either the air loop relief or outdoor air node
            string regenerationInletNode = hxObject["Exhaust Air Outlet Node Name"].Value;
            
            // Load all desiccant dehumidifer related objects
            string scheduleName = airLoopName + " Desiccant unit schedule";
            string dehumidifierSchedule = String.Format(dehumidifierScheduleBoilerplate, scheduleName);
            string dehumidiferFanName = airLoopName + " Desiccant Fan";

            string dehumidifierFan = String.Format(dehumidifierFanBoilerplate, dehumidiferFanName, scheduleName, regenerationInletNode);
            string dehumidifierSpm = String.Format(dehumidifierSpmBoilerplate, airLoopName, processOutletNode);

            string dehumidifierCoilName = airLoopName + " Desiccant Heating Coil";
            string dehumidifierCoil = String.Format(dehumidifierCoilBoilerplate, dehumidifierCoilName, scheduleName);
            string dehumidifierCoilType = "Coil:Heating:Electric";

            string dehumidifier = String.Format(
                dehumidifierBoilerplate,
                desiccantDehumidifierName, 
                scheduleName,
                processInletNode,
                processOutletNode,
                dehumidifierCoilType,
                regenerationInletNode, 
                dehumidifierCoilName, 
                dehumidiferFanName
                );

            reader.Load(dehumidifier);
            reader.Load(dehumidifierSchedule);
            reader.Load(dehumidifierFan);
            reader.Load(dehumidifierCoil);
            reader.Load(dehumidifierSpm);
        }


        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            AddDesiccantDehumifierToOaSystem(idfReader, "Air loop", 5);

            idfReader.Save();
        }
    }
}