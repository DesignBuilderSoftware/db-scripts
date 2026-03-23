/*
Add EvaporativeCooler:Direct:ResearchSpecial to an air loop branch.

Purpose:
This DesignBuilder C# script modifies the IDF by inserting an EvaporativeCooler:Direct:ResearchSpecial into a specified Branch for each target air loop name.

Main Steps:
1) Locate the target Branch object for each air loop (by name convention)
2) Break the existing node connection at a specific Branch field
3) Insert the evaporative cooler (type/name/inlet/outlet) into the Branch equipment list
4) Add supporting objects (availability schedule, setpoint schedule, setpoint manager, and output variables)

How to Use:

Configuration
- airLoopNames:
  Add one or more air loop base names, e.g. { "Air Loop 1", "Air Loop 2" }.
- Branch naming convention:
  For each airLoopName, the script expects a Branch named "<airLoopName> AHU Main Branch".
- Boilerplate template:
  The EvaporativeCooler object name is "<airLoopName> AHU Cooler" and related schedules/setpoint manager are created
  using that name as a prefix.

Prerequisites / Placeholders
Base model must contain a Branch object named "<airLoopName> AHU Main Branch" for each entry in airLoopNames.
These act as placeholders that allow the script to identify where to insert the new component.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. 
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Linq;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        // --------------------
        // USER CONFIGURATION 
        // --------------------
        // List of target air loop base names to modify.
        private readonly string[] targetAirLoopNames = new string[] { "Air Loop" };

        // IDF template that will be injected into the model for each target air loop.
        private readonly string coolerBoilerplateIdf = @"
EvaporativeCooler:Direct:ResearchSpecial,
  {0},                          !- Name
  {0} Availability,             !- Availability Schedule Name
  0.7 ,                         !- Cooler Design Effectiveness
  ,                             !- Effectiveness Flow Ratio Modifier Curve Name
  autosize,                     !- Primary Air Design Flow Rate m3/s
  30.0 ,                        !- Recirculating Water Pump Design Power
  ,                             !- Water Pump Power Sizing Factor
  ,                             !- Water Pump Power Modifier Curve Name
  {1},                          !- Air Inlet Node Name
  {2},                          !- Air Outlet Node Name
  {2},                          !- Sensor Node Name
  ,                             !- Water Supply Storage Tank Name
  0.0,                          !- Drift Loss Fraction
  3;                            !- Blowdown Concentration Ratio

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

Output:Variable, {0}, Evaporative Cooler Water Volume, Hourly;
Output:Variable, {0}, Evaporative Cooler Electricity Rate, Hourly;
Output:Variable, {0}, Evaporative Cooler Wet Bulb Effectiveness, Hourly;

Output:Variable, {2}, System Node Temperature, Hourly;";

        private IdfObject FindObject(IdfReader reader, string objectType, string objectName)
        {
            return reader[objectType].First(o => o[0] == objectName);
        }

        public override void BeforeEnergySimulation()
        {
            IdfReader reader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath
            );

            foreach (string airLoopName in targetAirLoopNames)
            {
                // Expected Branch naming convention in the IDF (must match exactly).
                string branchName = airLoopName + " AHU Main Branch";
                IdfObject branch = FindObject(reader, "Branch", branchName);

                // New evaporative cooler object name for this air loop.
                string coolerName = airLoopName + " AHU Cooler";

                // The script assumes branch[4] is the node currently connected where you want to insert the cooler.
                string coolerInletNode = branch[4].Value;

                // Create a new downstream node name to become the branch node after inserting the cooler.
                string coolerOutletNode = airLoopName + " AHU Cooler Outlet Node";

                // Rewire the branch so the cooler outlet becomes the node used at this position.
                branch[4].Value = coolerOutletNode;

                // Insert the cooler into the Branch equipment list
                string[] coolerBranchFields = new string[]
                {
                    "EvaporativeCooler:Direct:ResearchSpecial",
                    coolerName,
                    coolerInletNode,
                    coolerOutletNode
                };
                branch.InsertFields(2, coolerBranchFields);

                string coolerObjectsIdfText = String.Format(coolerBoilerplateIdf, coolerName, coolerInletNode, coolerOutletNode);
                reader.Load(coolerObjectsIdfText);
            }

            reader.Save();
        }
    }
}