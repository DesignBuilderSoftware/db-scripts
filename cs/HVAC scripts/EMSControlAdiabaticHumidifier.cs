/*
Adds EMS-based minimum humidity ratio setpoint control to one or more air loops

Purpose:
This DesignBuilder C# script injects an EnergyPlus Energy Management System (EMS) routine into the
IDF that dynamically overrides a dehumidification setpoint schedule for a given air loop. For each
configured air loop it reads the mixed-air node humidity ratio and enthalpy together with the supply
outlet minimum humidity ratio setpoint, and computes a dry-bulb temperature setpoint (via @TdbFnHW)
that will achieve the required minimum humidity ratio. When no dehumidification is needed the schedule
is driven to a high "off" value (99) so it has no effect on control.

Main Steps:
For each configured air loop:
1) Builds EnergyManagementSystem:Sensor objects for the supply outlet minimum humidity ratio,
   mixed-air outlet humidity ratio, and mixed-air outlet enthalpy
2) Builds an EnergyManagementSystem:Actuator on the target Schedule:Compact, a ProgramCallingManager,
   a Program implementing the control logic, and a debug Output:Variable
3) Loads the assembled EMS objects into the IDF via the IdfReader
Saves the modified IDF before the EnergyPlus simulation starts

How to Use:

Configuration
- ----------------------------
  USER CONFIGURATION SECTION
  ----------------------------
- In BeforeEnergySimulation(), add one ApplyTemperatureControl(...) call per air loop. For each call set:
  - airLoopName: Name of the air loop as it appears in the IDF (e.g. "AHU-5"). This is also used to
    build the node names "<airLoopName> Supply Side Outlet 1" and "<airLoopName> AHU Mixed Air Outlet".
  - setpointScheduleName: Name of the existing Schedule:Compact that will be actuated by the EMS
    (e.g. "AHU_5_AH_SETPOINT").
- Internal EMS object names are derived automatically by replacing spaces and hyphens with underscores,
  so they are valid EnergyPlus identifiers.

Prerequisites / Placeholders
- The base model must already contain the named air loop.
- The scripts requires a humidity minimum setpoint manager on the AHU outlet

Notes:
- A debug Output:Variable for the actuated schedule value is added at Timestep frequency.


DISCLAIMER: This script is provided as-is without warranty. 
DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. 
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System.Runtime;
using System.Linq;
using System;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        private IdfReader Reader;

        public override void BeforeEnergySimulation()
        {
            Reader = new IdfReader(
              ApiEnvironment.EnergyPlusInputIdfPath,
              ApiEnvironment.EnergyPlusInputIddPath);

            // ----------------------------
            // USER CONFIGURATION SECTION
            // ----------------------------
            // Define air loops to control (repeat this call for each air loop).
            // Args: airLoopName (IDF air loop name), setpointScheduleName (existing Schedule:Compact to actuate)
            ApplyTemperatureControl("Air Loop", "AHU_SETPOINT");

            Reader.Save();
        }

        public void ApplyTemperatureControl(string airLoopName, string setpointScheduleName)
        {
            // Derive valid EMS identifiers
            string emsAirLoopName = airLoopName.Replace(" ", "_").Replace("-", "_");
            string emsScheduleName = setpointScheduleName.Replace(" ", "_").Replace("-", "_");

            // EMS object template:
            //   {0} = airLoopName (used for node names)
            //   {1} = emsAirLoopName (EMS-safe air loop identifier)
            //   {2} = setpointScheduleName (Schedule:Compact to actuate / output)
            //   {3} = emsScheduleName (EMS-safe actuator name)
            string template = @"
! Read humidity indicators
EnergyManagementSystem:Sensor,
   {1}_Setpoint_Minimum_Humidity_Ratio,
   {0} Supply Side Outlet 1,
   System Node Setpoint Minimum Humidity Ratio;
EnergyManagementSystem:Sensor,
   {1}_Humidity_Ratio,
   {0} AHU Mixed Air Outlet,
   System Node Humidity Ratio;
EnergyManagementSystem:Sensor,
   {1}_Enthalpy,
   {0} AHU Mixed Air Outlet,
   System Node Enthalpy;
! Actuate schedules
EnergyManagementSystem:Actuator,
   {3},
   {2},
   Schedule:Compact,
   Schedule Value;
! Define debug outputs
Output:Variable, {2}, Schedule Value, Timestep;
EnergyManagementSystem:ProgramCallingManager,
   {1}_Caller,
   InsideHVACSystemIterationLoop,
   {1}_Control;
EnergyManagementSystem:Program,
   {1}_Control,
   IF {1}_Humidity_Ratio < {1}_Setpoint_Minimum_Humidity_Ratio,
      SET {3} = @TdbFnHW {1}_Enthalpy {1}_Setpoint_Minimum_Humidity_Ratio,
   ELSE,
      SET {3} = 99,
   ENDIF;
";

            string content = String.Format(template, airLoopName, emsAirLoopName, setpointScheduleName, emsScheduleName);
            Reader.Load(content);
        }

        public IdfObject FindObject(string objectType, string objectName)
        {
            try
            {
                return Reader[objectType].First(c => c[0] == objectName);
            }
            catch (Exception)
            {
                throw new MissingFieldException(String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }
    }
}