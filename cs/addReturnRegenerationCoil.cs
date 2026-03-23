/*
Add Electric Heating Coil to an Air Loop Branch (Return Air Stream)

Purpose:
This DesignBuilder C# script inserts a Coil:Heating:Electric into the air loop branch for selected air loops by editing the EnergyPlus IDF before simulation.

Main Steps:
For each target air loop name:
1) Find the corresponding Branch object (following naming convention "<Air Loop Name> AHU Main Branch")
2) Identify the existing downstream node currently referenced by the branch
3) Insert a new coil component into the branch equipment list
4) Add required supporting objects (coil, availability schedule, setpoint schedule, scheduled setpoint manager)
5) Add a couple of Output:Variable requests

How to Use:

Configuration
- targetAirLoopNames:
  - List of air loop names to modify (exact match required).
- coilAndControlsIdfTemplate:
  - IDF template injected for each air loop, including:
    - Coil:Heating:Electric
    - Schedule:Compact (Availability)
    - Schedule:Compact (Setpoint)
    - SetpointManager:Scheduled
    - Output:Variable entries

Prerequisites / Placeholders
- Each target air loop must already contain a Branch object named "<Air Loop Name> AHU Main Branch".

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. 
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Linq;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class AddElectricHeatingCoilToAirLoopBranch : ScriptBase, IScript
    {
        // Configure which air loops to modify (names must match the model exactly)
        private readonly string[] targetAirLoopNames = new string[] { "Air Loop" };

        // IDF template injected per air loop:
        // {0} = coil name
        // {1} = coil inlet node name
        // {2} = coil outlet node name (also used as setpoint node)
        private readonly string coilAndControlsIdfTemplate = @"
Coil:Heating:Electric,
  {0},                          !- Name
  {0} Availability,             !- Availability Schedule Name
  1.00,                         !- Efficiency (%)
  autosize,                     !- Nominal capacity (W)
  {1},                          !- Air Inlet Node Name
  {2},                          !- Air Outlet Node Name
  {2};                          !- Sensor Node Name

Schedule:Compact,
   {0} Availability,            ! Name
   Any Number,                  ! Type
   Through: 12/31,              ! Type
   For: AllDays,                ! All days in year
   Until: 24:00,                ! All hours in day
   1;

Schedule:Compact,
   {0} Setpoint,                ! Name
   Any Number,                  ! Type
   Through: 12/31,              ! Type
   For: AllDays,                ! All days in year
   Until: 24:00,                ! All hours in day
   18;

SetpointManager:Scheduled,
  {0} Setpoint Manager,         !- Name
  Temperature,                  !- Control Variable
  {0} Setpoint,                 !- Schedule Name
  {2};                          !- Setpoint Node or NodeList Name

Output:Variable, {0}, Heating Coil Electricity Rate, Hourly;
Output:Variable, {2}, System Node Temperature, Hourly;
";

        // Helper: Find a single IDF object by type + name (name is typically field 0)
        private IdfObject FindObjectByName(IdfReader idf, string objectType, string objectName)
        {
            return idf[objectType].First(o => o[0] == objectName);
        }

        public override void BeforeEnergySimulation()
        {
            IdfReader idf = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath
            );

            foreach (string airLoopName in targetAirLoopNames)
            {
                string mainBranchName = airLoopName + " AHU Main Branch";
                IdfObject mainBranch = FindObjectByName(idf, "Branch", mainBranchName);

                string coilName = airLoopName + " AHU Regeneration Coil";
                // Assumption: branch[4] is the node currently used downstream in the branch definition.
                // Preserve it as the coil inlet node, then replace branch[4] with the new coil outlet node.
                string coilInletNode = mainBranch[4].Value;
                string coilOutletNode = airLoopName + " AHU Regeneration Coil Outlet Node";

                mainBranch[4].Value = coilOutletNode;

                // Insert the new coil into the branch equipment list.
                string[] newBranchFields = new string[]
                {
                    "Coil:Heating:Electric",
                    coilName,
                    coilInletNode,
                    coilOutletNode
                };
                mainBranch.InsertFields(2, newBranchFields);

                string idfTextToLoad = String.Format(coilAndControlsIdfTemplate, coilName, coilInletNode, coilOutletNode);
                idf.Load(idfTextToLoad);
            }

            idf.Save();
        }
    }
}