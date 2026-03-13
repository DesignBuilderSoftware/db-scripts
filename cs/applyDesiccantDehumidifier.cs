/*
Add Desiccant Dehumidifier to an Air Handling Unit (AHU) Supply Path

This DesignBuilder C# script inserts an EnergyPlus Dehumidifier:Desiccant:NoFans component into an existing air loop supply branch
at a specified position, then rewires the downstream component inlet node so airflow passes through the dehumidifier.

Purpose

1) Locate the AHU main supply Branch for a given air loop
2) Insert Dehumidifier:Desiccant:NoFans into the Branch component list at a given position
3) Create required supporting objects:
       * Schedule:Compact (availability)
       * Fan:VariableVolume (regeneration fan)
       * Regeneration heating coil (Coil:Heating:Fuel or Coil:Heating:Water)
       * SetpointManager:MultiZone:Humidity:Maximum
       * OutdoorAir:Node (regeneration air inlet node)
4) Update the following component’s inlet node (and related inlet fields for certain systems)
     to match the dehumidifier process outlet node, ensuring correct node continuity.

How to Use

Configuration
AddDesiccantDehumidifier parameters:
   - idfReader: IdfReader instance for the current simulation IDF/IDD
   - airLoopName: Name of the target EnergyPlus AirLoopHVAC (used to derive object names)
   - position: Component position in the AHU main Branch supply path (1-based)
   - regenerationCoilCapacity: Nominal regen coil capacity, in W
   - regenerationCoilType: RegenerationCoilType.Fuel or RegenerationCoilType.Water (default Fuel)
   - hwLoopName: Plant loop name (required when regenerationCoilType = Water)

Prerequisites (required placeholders)

This script expects an Air Loop to be in place in the model (airLoopName referenced in AddDesiccantDehumidifier())
A hot water water loop must be in place if using RegenerationCoilType.Water (hwLoopName referenced in AddDesiccantDehumidifier())

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class AddDesiccantDehumidifierToAhu : ScriptBase, IScript
    {
        // IDF text templates for objects created by this script.
        private readonly string dehumidifierIdfTemplate = @"
Dehumidifier:Desiccant:NoFans,
    {0} Desiccant Dehumidifier,                                 !- Name
    {0} Dehumidifier Schedule,                                  !- Availability Schedule Name
    {1},                                                        !- Process Air Inlet Node Name
    {2},                                                        !- Process Air Outlet Node Name
    {0} Desiccant Heating Coil Air Outlet Node,                 !- Regeneration Air Inlet Node Name
    {0} Outside Air Inlet Node 2,                               !- Regeneration Fan Inlet Node Name
    SystemNodeMaximumHumidityRatioSetpoint,                     !- Control Type
    0.007,                                                      !- Leaving Maximum Humidity Ratio Setpoint [kgWater/kgDryAir]
    1.5,                                                        !- Nominal Process Air Flow Rate [m3/s]
    2.5,                                                        !- Nominal Process Air Velocity [m/s]
    50,                                                         !- Rotor Power [W]
    {3},                                                        !- Regeneration Coil Object Type
    {0} Desiccant Regen Coil,                                   !- Regeneration Coil Name
    Fan:VariableVolume,                                         !- Regeneration Fan Object Type
    {0} Desiccant Regen Fan,                                    !- Regeneration Fan Name
    DEFAULT;                                                    !- Performance Model Type

OutdoorAir:Node,
    {0} Outside Air Inlet Node 2;";

        private readonly string regenFanIdfTemplate = @"
Fan:VariableVolume,
    {0} Desiccant Regen Fan,                                    !- Name
    {0} Dehumidifier Schedule,                                  !- Availability Schedule Name
    0.7,                                                        !- Fan Total Efficiency
    600.0,                                                      !- Pressure Rise [Pa]
    Autosize,                                                   !- Maximum Flow Rate [m3/s]
    FixedFlowRate,                                              !- Fan Power Minimum Flow Rate Input Method
    ,                                                          !- Fan Power Minimum Flow Fraction
    0.0,                                                        !- Fan Power Minimum Air Flow Rate [m3/s]
    0.9,                                                        !- Motor Efficiency
    1.0,                                                        !- Motor In Airstream Fraction
    0,                                                          !- Fan Power Coefficient 1
    1,                                                          !- Fan Power Coefficient 2
    0,                                                          !- Fan Power Coefficient 3
    0,                                                          !- Fan Power Coefficient 4
    0,                                                          !- Fan Power Coefficient 5
    {0} Outside Air Inlet Node 2,                               !- Air Inlet Node Name
    {0} Regen Fan Outlet Node;                                  !- Air Outlet Node Name";

        private readonly string regenWaterCoilAndBranchIdfTemplate = @"
Coil:Heating:Water,
    {0} Desiccant Regen Coil,                                   !- Name
    {0} Dehumidifier Schedule,                                  !- Availability Schedule Name
    autosize,                                                   !- U-Factor Times Area Value [W/K]
    autosize,                                                   !- Maximum Water Flow Rate [m3/s]
    {0} Desiccant Heating Coil Water Inlet Node,                !- Water Inlet Node Name
    {0} Desiccant Heating Coil Water Outlet Node,               !- Water Outlet Node Name
    {0} Regen Fan Outlet Node,                                  !- Air Inlet Node Name
    {0} Desiccant Heating Coil Air Outlet Node,                 !- Air Outlet Node Name
    NominalCapacity,                                            !- Performance Input Method
    {1},                                                        !- Rated Capacity [W]
    80,                                                         !- Rated Inlet Water Temperature [C]
    16,                                                         !- Rated Inlet Air Temperature [C]
    70,                                                         !- Rated Outlet Water Temperature [C]
    35,                                                         !- Rated Outlet Air Temperature [C]
    0.50;                                                       !- Rated Ratio for Air and Water Convection

Branch,
    {0} Desiccant Regen Coil Branch,                            !- Name
    ,                                                          !- Pressure Drop Curve Name
    Coil:Heating:Water,                                         !- Component 1 Object Type
    {0} Desiccant Regen Coil,                                   !- Component 1 Name
    {0} Desiccant Heating Coil Water Inlet Node,                !- Component 1 Inlet Node Name
    {0} Desiccant Heating Coil Water Outlet Node;               !- Component 1 Outlet Node Name";

        private readonly string regenFuelCoilIdfTemplate = @"
Coil:Heating:Fuel,
    {0} Desiccant Regen Coil,                                   !- Name
    {0} Dehumidifier Schedule,                                  !- Availability Schedule Name
    NaturalGas,                                                 !- Fuel Type
    0.80,                                                       !- Burner Efficiency
    {1},                                                        !- Nominal Capacity [W]
    {0} Regen Fan Outlet Node,                                  !- Air Inlet Node Name
    {0} Desiccant Heating Coil Air Outlet Node;                 !- Air Outlet Node Name";

        private readonly string availabilityScheduleIdfTemplate = @"
Schedule:Compact,
    {0} Dehumidifier Schedule,                                  !- Name
    Any Number,                                                 !- Schedule Type Limits Name
    Through: 12/31,
    For: AllDays,
    Until: 24:00,
    1;";

        private readonly string humiditySpmIdfTemplate = @"
SetpointManager:MultiZone:Humidity:Maximum,
    {0} Humidity Setpoint Manager,                              !- Name
    {0},                                                        !- HVAC Air Loop Name
    .005,                                                       !- Minimum Setpoint Humidity Ratio [kg/kg]
    .012,                                                       !- Maximum Setpoint Humidity Ratio [kg/kg]
    {1};                                                        !- Setpoint Node List Name (or Node Name, depending on model)";

        public enum RegenerationCoilType
        {
            Water,
            Fuel
        }

        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // ----------------------------
            // USER CONFIGURATION SECTION
            // ----------------------------
            // This example usage inserts a dehumidifier into "Air Loop" main branch at position 3, using a fuel regeneration coil of 20 kW capacity.
            AddDesiccantDehumidifier(
                idfReader,
                airLoopName: "Air Loop",
                position: 3,
                regenerationCoilCapacity: 20000,
                regenerationCoilType: RegenerationCoilType.Fuel);

            idfReader.Save();
        }

        public void AddDesiccantDehumidifier(
            IdfReader idfReader,
            string airLoopName,
            int position,
            int regenerationCoilCapacity,
            RegenerationCoilType regenerationCoilType = RegenerationCoilType.Fuel,
            string hwLoopName = "")
        {
            const string dehumidifierObjectType = "Dehumidifier:Desiccant:NoFans";
            string mainSupplyBranchName = airLoopName + " AHU Main Branch";
            IdfObject mainSupplyBranch = FindObject(idfReader, "Branch", mainSupplyBranchName);

            // Branch objects store components in repeating groups of 4 fields (ObjectType, ObjectName, InletNode, OutletNode)
            // The first component group starts at field index 2 (0-based).
            int dehumidifierBranchFieldIndex = 2 + (position - 1) * 4;

            // Determine process air inlet/outlet nodes for the new dehumidifier.
            // The inlet is the node that previously fed the component at 'position'.
            string processAirInletNode = mainSupplyBranch[dehumidifierBranchFieldIndex - 1].Value;
            string processAirOutletNode = airLoopName + " Desiccant Process Outlet Node";

            // Identify the component currently at this position to rewire its inlet to the new outlet.
            string downstreamComponentType = mainSupplyBranch[dehumidifierBranchFieldIndex].Value;
            string downstreamComponentName = mainSupplyBranch[dehumidifierBranchFieldIndex + 1].Value;
            IdfObject downstreamComponent = FindObject(idfReader, downstreamComponentType, downstreamComponentName);

            // Component node field names may vary depending on the component type
            if (downstreamComponentType.Equals("CoilSystem:Heating:DX", StringComparison.OrdinalIgnoreCase))
            {
                IdfObject heatingCoil = FindObject(
                    idfReader,
                    downstreamComponent["Heating Coil Object Type"].Value,
                    downstreamComponent["Heating Coil Name"].Value);

                heatingCoil["Air Inlet Node Name"].Value = processAirOutletNode;
            }
            else if (downstreamComponentType.Equals("CoilSystem:Cooling:DX", StringComparison.OrdinalIgnoreCase))
            {
                IdfObject coolingCoil = FindObject(
                    idfReader,
                    downstreamComponent["Cooling Coil Object Type"].Value,
                    downstreamComponent["Cooling Coil Name"].Value);

                coolingCoil["Air Inlet Node Name"].Value = processAirOutletNode;
                downstreamComponent["DX Cooling Coil System Inlet Node Name"].Value = processAirOutletNode;
            }
            else
            {
                downstreamComponent["Air Inlet Node Name"].Value = processAirOutletNode;
            }

            // Update the Branch entry for the downstream component so its inlet node matches the new node.
            mainSupplyBranch[dehumidifierBranchFieldIndex + 2].Value = processAirOutletNode;

            // Insert dehumidifier into the AHU main supply branch component list.
            string[] newBranchFields = new string[]
            {
                dehumidifierObjectType,
                airLoopName + " Desiccant Dehumidifier",
                processAirInletNode,
                processAirOutletNode
            };
            mainSupplyBranch.InsertFields(dehumidifierBranchFieldIndex, newBranchFields);

            // Build IDF text for all supporting objects (schedule, fan, coil, setpoint manager, OA node).
            string scheduleIdfText = string.Format(availabilityScheduleIdfTemplate, airLoopName);
            string regenFanIdfText = string.Format(regenFanIdfTemplate, airLoopName);
            string humiditySpmIdfText = string.Format(humiditySpmIdfTemplate, airLoopName, processAirOutletNode);

            string regenCoilIdfText;
            string regenCoilObjectType;

            switch (regenerationCoilType)
            {
                case RegenerationCoilType.Water:
                    regenCoilIdfText = string.Format(regenWaterCoilAndBranchIdfTemplate, airLoopName, regenerationCoilCapacity);
                    regenCoilObjectType = "Coil:Heating:Water";

                    // For water regen coils, the script also expects to connect the new coil branch
                    // into the HW loop demand side branch list/splitter/mixer (by naming convention).
                    AddDemandSideBranchToPlantLoop(idfReader, hwLoopName, airLoopName + " Desiccant Regen Coil Branch");
                    break;

                case RegenerationCoilType.Fuel:
                    regenCoilIdfText = string.Format(regenFuelCoilIdfTemplate, airLoopName, regenerationCoilCapacity);
                    regenCoilObjectType = "Coil:Heating:Fuel";
                    break;

                default:
                    throw new ArgumentException("Invalid regeneration coil type specified.");
            }

            string dehumidifierIdfText = string.Format(
                dehumidifierIdfTemplate,
                airLoopName,
                processAirInletNode,
                processAirOutletNode,
                regenCoilObjectType);

            idfReader.Load(dehumidifierIdfText);
            idfReader.Load(scheduleIdfText);
            idfReader.Load(regenFanIdfText);
            idfReader.Load(regenCoilIdfText);
            idfReader.Load(humiditySpmIdfText);
        }

        public IdfObject FindObject(IdfReader reader, string objectType, string objectName)
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

        private void AddDemandSideBranchToPlantLoop(IdfReader reader, string loopName, string branchName)
        {
            IdfObject branchList = FindObject(reader, "BranchList", loopName + " Demand Side Branches");
            branchList.InsertField(branchList.Count - 1, branchName);

            IdfObject splitter = FindObject(reader, "Connector:Splitter", loopName + " Demand Splitter");
            splitter.InsertField(splitter.Count - 1, branchName);

            IdfObject mixer = FindObject(reader, "Connector:Mixer", loopName + " Demand Mixer");
            mixer.InsertField(mixer.Count - 1, branchName);
        }
    }
}