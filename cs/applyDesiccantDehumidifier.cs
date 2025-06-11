/*
Add Dehumidifier:Desiccant:NoFans object into air handling unit.

The AddDesiccantDehumidifier method requires the following parameters:
    - air loop name
    - position (component position in the supply path)
    - regeneration coil type (Water, Fuel)
    - hw loop name (if using water coil)

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
    {0} Desiccant Dehumidifier,!- Name
    {0} Dehumidifier Schedule, !- Availability Schedule Name
    {1},                     !- Process Air Inlet Node Name
    {2},                     !- Process Air Outlet Node Name
    {0} Desiccant Heating Coil Air Outlet Node,     !- Regeneration Air Inlet Node Name
    {0} Outside Air Inlet Node 2,!- Regeneration Fan Inlet Node Name
    SystemNodeMaximumHumidityRatioSetpoint,  !- Control Type alt LeavingMaximumHumidityRatioSetpoint
    0.007,                   !- Leaving Maximum Humidity Ratio Setpoint kgWater/kgDryAir
    1.5,                     !- Nominal Process Air Flow Rate m3/s
    2.5,                     !- Nominal Process Air Velocity m/s
    50,                      !- Rotor Power W
    {3},                     !- Regeneration Coil Object Type
    {0} Desiccant Regen Coil,!- Regeneration Coil Name
    Fan:VariableVolume,      !- Regeneration Fan Object Type
    {0} Desiccant Regen Fan, !- Regeneration Fan Name
    DEFAULT;                 !- Performance Model Type

    OutdoorAir:Node,{0} Outside Air Inlet Node 2;";

        string dehumidifierFanBoilerplate = @"  Fan:VariableVolume,
    {0} Desiccant Regen Fan,     !- Name
    {0} Dehumidifier Schedule,   !- Availability Schedule Name
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
    {0} Outside Air Inlet Node 2,!- Air Inlet Node Name
    {0} Regen Fan Outlet Node;   !- Air Outlet Node Name";

        string dehumidifierWaterCoilBoilerplate = @"  Coil:Heating:Water,
  {0} Desiccant Regen Coil,                                       ! - Component name
  {0} Dehumidifier Schedule,                                      ! - Availability schedule
  autosize,                                                       ! - U-factor times area value of coil (W/K)
  autosize,                                                       ! - Max water flow rate of coil (m3/s)
  {0} Desiccant Heating Coil Water Inlet Node,                    ! - Water inlet node name
  {0} Desiccant Heating Coil Water Outlet Node,                   ! - Water outlet node name
  {0} Regen Fan Outlet Node,                                      ! - Air inlet node name
  {0} Desiccant Heating Coil Air Outlet Node,                     ! - Air outlet node name
  NominalCapacity,                                                ! - Coil performance input method
  {1},                                                            ! - Rated capacity (W)
  80,                                                             ! - Rated inlet water temperature (C)
  16,                                                             ! - Rated inlet air temperature (C)
  70,                                                             ! - Rated outlet water temperature (C)
  35,                                                             ! - Rated outlet air temperature (C)
  0.50;                                                           ! - Rated ratio for air and water convection

Branch,
  {0} Desiccant Regen Coil Branch,                                ! - Branch name
  ,                                                               ! - Pressure drop curve name
  Coil:Heating:Water,                                             ! - Component 1 object type
  {0} Desiccant Regen Coil,                                       ! - Component 1 name
  {0} Desiccant Heating Coil Water Inlet Node,                    ! - Component 1 inlet node name
  {0} Desiccant Heating Coil Water Outlet Node;                   ! - Component 1 outlet node name";

        string dehumidifierCoilBoilerplate = @"  Coil:Heating:Fuel,
    {0} Desiccant Regen Coil,    !- Name
    {0} Dehumidifier Schedule,   !- Availability Schedule Name
    NaturalGas,                  !- Fuel Type
    0.80,                        !- Burner Efficiency
    {1},                         !- Nominal Capacity W
    {0} Regen Fan Outlet Node,   !- Air Inlet Node Name
    {0} Desiccant Heating Coil Air Outlet Node;     !- Air Outlet Node Name";

        string dehumidifierScheduleBoilerplate = @"  Schedule:Compact, 
   {0} Dehumidifier Schedule,! Name
   Any Number,               ! Type
   Through: 12/31,           ! Type
   For: AllDays,             ! All days in year
   Until: 24:00,             ! All hours in day
   1;";

        string dehumidifierSpmBoilerplate = @"  SetpointManager:MultiZone:Humidity:Maximum,
   {0} Humidity Setpoint Manager,                       ! - Component name
   {0},                                                 ! - HVAC air loop name
   .005,                                                ! - Minimum setpoint humidity ratio (kg/kg)
   .012,                                                ! - Maximum setpoint humidity ratio (kg/kg)
   {1};                                                 ! - Setpoint node list";

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

            AddDesiccantDehumidifier(idfReader, "Air Loop", 3, 20000, RegenerationCoilType.Fuel);

            idfReader.Save();
        }

        public void AddDesiccantDehumidifier(
            IdfReader idfReader,
            string airLoopName, 
            int position,
            int regenerationCoilCapacity,
            RegenerationCoilType regenerationCoilType = RegenerationCoilType.Fuel,
            string hwLoopName = ""
            )
        {
            string desiccantDehumidifierComponent = "Dehumidifier:Desiccant:NoFans";
            string ahuMainBranchName = airLoopName + " AHU Main Branch";

            IdfObject ahuMainBranch = FindObject(idfReader, "Branch", ahuMainBranchName);

            // calculate index of the field set in the main branch
            int dehumidfierBranchIndex = 2 + (position - 1) * 4;
 
            // update component nodes
            string processInletNode = ahuMainBranch[dehumidfierBranchIndex - 1].Value;
            string processOutletNode = airLoopName + " Desiccant Process Outlet Node";

            string nextComponentType = ahuMainBranch[dehumidfierBranchIndex].Value;
            string nextComponentName = ahuMainBranch[dehumidfierBranchIndex + 1].Value;

            IdfObject nextComponent = FindObject(idfReader, nextComponentType, nextComponentName);

            // component node field names may vary depending on the component type
            if (nextComponentType.Equals("CoilSystem:Heating:DX", StringComparison.OrdinalIgnoreCase))
            {
                IdfObject coil = FindObject(idfReader, nextComponent["Heating Coil Object Type"].Value, nextComponent["Heating Coil Name"].Value);
                coil["Air Inlet Node Name"].Value = processOutletNode;
            } 
            else if (nextComponentType.Equals("CoilSystem:Cooling:DX", StringComparison.OrdinalIgnoreCase))
            {
                IdfObject coil = FindObject(idfReader, nextComponent["Cooling Coil Object Type"].Value, nextComponent["Cooling Coil Name"].Value);
                coil["Air Inlet Node Name"].Value = processOutletNode;
                nextComponent["DX Cooling Coil System Inlet Node Name"].Value = processOutletNode;
            }
            else
            {
                nextComponent["Air Inlet Node Name"].Value = processOutletNode;
            }
            ahuMainBranch[dehumidfierBranchIndex + 2].Value = processOutletNode;

            // insert dehumidifier component into AHU main branch
            string[] newBranchFields = new string[] { desiccantDehumidifierComponent, airLoopName + " Desiccant Dehumidifier", processInletNode, processOutletNode };
            ahuMainBranch.InsertFields(dehumidfierBranchIndex, newBranchFields);

            // create object idf content
            string dehumidifierSchedule = String.Format(dehumidifierScheduleBoilerplate, airLoopName);
            string dehumidifierFan = String.Format(dehumidifierFanBoilerplate, airLoopName);
            string dehumidifierSpm = String.Format(dehumidifierSpmBoilerplate, airLoopName, processOutletNode);
            string dehumidifierCoil = "";
            string dehumidifierCoilType = "";

            switch (regenerationCoilType)
            {
                case RegenerationCoilType.Water:
                    dehumidifierCoil = String.Format(dehumidifierWaterCoilBoilerplate, airLoopName, regenerationCoilCapacity);
                    dehumidifierCoilType = "Coil:Heating:Water";
                    AddBranch(idfReader, hwLoopName, airLoopName + " Desiccant Regen Coil Branch");
                    break;
                case RegenerationCoilType.Fuel:
                    dehumidifierCoil = String.Format(dehumidifierCoilBoilerplate, airLoopName, regenerationCoilCapacity);
                    dehumidifierCoilType = "Coil:Heating:Fuel";
                    break;
                default:
                    throw new ArgumentException("Invalid regeneration coil type specified.");
            }

            string dehumidifier = String.Format(dehumidifierBoilerplate, airLoopName, processInletNode, processOutletNode, dehumidifierCoilType);

            idfReader.Load(dehumidifier);
            idfReader.Load(dehumidifierSchedule);
            idfReader.Load(dehumidifierFan);
            idfReader.Load(dehumidifierCoil);
            idfReader.Load(dehumidifierSpm);
        }

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

        private void AddBranch(IdfReader reader, string loopName, string branchName)
        {
            IdfObject branchList = FindObject(reader, "branchList", loopName + " Demand Side Branches");
            branchList.InsertField(branchList.Count - 1, branchName);

            IdfObject splitter = FindObject(reader, "Connector:Splitter", loopName + " Demand Splitter");
            splitter.InsertField(splitter.Count - 1, branchName);

            IdfObject mixer = FindObject(reader, "Connector:Mixer", loopName + " Demand Mixer");
            mixer.InsertField(mixer.Count - 1, branchName);
        }
    }
}