/*
Add Dehumidifier:Desiccant:NoFans object into air handling unit.

The component is added between OA mixer and the first component in the main branch
of loop named "DesiccantDehumidifier".

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
    Outside Air Inlet Node 2,!- Regeneration Fan Inlet Node Name
    SystemNodeMaximumHumidityRatioSetpoint,  !- Control Type alt LeavingMaximumHumidityRatioSetpoint
    0.007,                   !- Leaving Maximum Humidity Ratio Setpoint kgWater/kgDryAir
    1.5,                     !- Nominal Process Air Flow Rate m3/s
    2.5,                     !- Nominal Process Air Velocity m/s
    50,                      !- Rotor Power W
    {4},                     !- Regeneration Coil Object Type
    Desiccant Regen Coil,    !- Regeneration Coil Name
    Fan:VariableVolume,      !- Regeneration Fan Object Type
    Desiccant Regen Fan,     !- Regeneration Fan Name
    DEFAULT;                 !- Performance Model Type

    OutdoorAir:Node, Outside Air Inlet Node 2;";

        string dehumidifierFanBoilerplate = @"  Fan:VariableVolume,
    Desiccant Regen Fan,     !- Name
    {0},                     !- Availability Schedule Name
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
    Outside Air Inlet Node 2,!- Air Inlet Node Name
    Regen Fan Outlet Node;   !- Air Outlet Node Name";

        string dehumidifierWaterCoilBoilerplate = @"  Coil:Heating:Water,
  Desiccant Regen Coil,                                           ! - Component name
  {0},                                                            ! - Availability schedule
  autosize,                                                       ! - U-factor times area value of coil (W/K)
  autosize,                                                       ! - Max water flow rate of coil (m3/s)
  Desiccant Heating Coil Water Inlet Node,                        ! - Water inlet node name
  Desiccant Heating Coil Water Outlet Node,                       ! - Water outlet node name
  Regen Fan Outlet Node,                                          ! - Air inlet node name
  Desiccant Heating Coil Air Outlet Node,                         ! - Air outlet node name
  NominalCapacity,                                                ! - Coil performance input method
  200000,                                                         ! - Rated capacity (W)
  80,                                                             ! - Rated inlet water temperature (C)
  16,                                                             ! - Rated inlet air temperature (C)
  70,                                                             ! - Rated outlet water temperature (C)
  35,                                                             ! - Rated outlet air temperature (C)
  0.50;                                                           ! - Rated ratio for air and water convection

Branch,
  {1},                                                            ! - Branch name
  ,                                                               ! - Pressure drop curve name
  Coil:Heating:Water,                                             ! - Component 1 object type
  Desiccant Regen Coil,                                           ! - Component 1 name
  Desiccant Heating Coil Water Inlet Node,                        ! - Component 1 inlet node name
  Desiccant Heating Coil Water Outlet Node;                       ! - Component 1 outlet node name";


        string dehumidifierCoilBoilerplate = @"  Coil:Heating:Fuel,
    Desiccant Regen Coil,    !- Name
    {0},                     !- Availability Schedule Name
    NaturalGas,              !- Fuel Type
    0.80,                    !- Burner Efficiency
    200000,                  !- Nominal Capacity W
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

        private void AddBranch(IdfReader reader, string loopName, string branchName)
        {
            IdfObject branchList = FindObject(reader, "branchList", loopName + " Demand Side Branches");
            branchList.InsertField(branchList.Count - 1, branchName);

            IdfObject splitter = FindObject(reader, "Connector:Splitter", loopName + " Demand Splitter");
            splitter.InsertField(splitter.Count - 1, branchName);

            IdfObject mixer = FindObject(reader, "Connector:Mixer", loopName + " Demand Mixer");
            mixer.InsertField(mixer.Count - 1, branchName);
        }


        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            bool waterCoil = false;    // use Coil:Heating:Water when true
            string ahuName = "Air Loop";    // specify air loop name
            int dehumidfierBranchIndex = 10;    // place the component into the nth index in the AHU Main Branch

            string desiccantDehumidifierComponent = "Dehumidifier:Desiccant:NoFans";
            string desiccantDehumidifierName = "Desiccant Unit";
            string ahuMainBranchName = ahuName + " AHU Main Branch";
            string hwLoopName = "HW Loop";
            string desiccantHwLoopBranchName = "Desiccant HW Branch";

            IdfObject ahuMainBranch = FindObject(idfReader, "Branch", ahuMainBranchName);


            string processInletNode = ahuMainBranch[dehumidfierBranchIndex - 1].Value;
            string processOutletNode = "Desiccant Process Outlet Node";

            string nextComponentType = ahuMainBranch[dehumidfierBranchIndex].Value;
            string nextComponentName = ahuMainBranch[dehumidfierBranchIndex + 1].Value;

            IdfObject nextComponent = FindObject(idfReader, nextComponentType, nextComponentName);
            nextComponent["Air Inlet Node Name"].Value = processOutletNode;
            ahuMainBranch[dehumidfierBranchIndex + 2].Value = processOutletNode;

            string[] newBranchFields = new string[] { desiccantDehumidifierComponent, desiccantDehumidifierName, processInletNode, processOutletNode };

            ahuMainBranch.InsertFields(dehumidfierBranchIndex, newBranchFields);

            string scheduleName = "Desiccant unit schedule";

            string dehumidifierSchedule = String.Format(dehumidifierScheduleBoilerplate, scheduleName);
            string dehumidifierFan = String.Format(dehumidifierFanBoilerplate, scheduleName);
            string dehumidifierSpm = String.Format(dehumidifierSpmBoilerplate, ahuName, processOutletNode);
            string dehumidifierCoil = "";
            string dehumidifierCoilType = "";

            if (waterCoil)
            {
                AddBranch(idfReader, hwLoopName, desiccantHwLoopBranchName);
                dehumidifierCoil = String.Format(dehumidifierWaterCoilBoilerplate, scheduleName, desiccantHwLoopBranchName);
                dehumidifierCoilType = "Coil:Heating:Water";
            }
            else
            {
                dehumidifierCoil = String.Format(dehumidifierCoilBoilerplate, scheduleName);
                dehumidifierCoilType = "Coil:Heating:Fuel";
            }

            string dehumidifier = String.Format(dehumidifierBoilerplate, desiccantDehumidifierName, scheduleName, processInletNode, processOutletNode, dehumidifierCoilType);

            idfReader.Load(dehumidifier);
            idfReader.Load(dehumidifierSchedule);
            idfReader.Load(dehumidifierFan);
            idfReader.Load(dehumidifierCoil);
            idfReader.Load(dehumidifierSpm);

            idfReader.Save();
        }
    }
}