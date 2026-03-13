/*
Zone Fan-Coil Night Cycle Operation (AvailabilityManager:NightCycle)

Purpose
This DesignBuilder C# script enables Night Cycle Operation control for zone fan-coil units
(ZoneHVAC:FourPipeFanCoil) by adding an AvailabilityManager:NightCycle and assignment list.

DesignBuilder automatically applies the unit availability schedule to the child fan.
This setup would not work well for night cycle as the manager overrides only the fan availability.

Purpose

1) Find all ZoneHVAC:FourPipeFanCoil objects in the IDF.
2) For each fan-coil unit:
   - Assign an AvailabilityManagerAssignmentList to the unit.
   - Force the unit Availability Schedule Name to "On 24/7" (so the availability manager can drive operation).
   - Read the fan’s Availability Schedule Name and use it in the AvailabilityManager:NightCycle object.
3) Load the generated IDF objects and save the modified IDF.

How to Use

Configuration
- Constant schedule name used to force unit availability ("On 24/7").
- Control zone name derivation:
  - The script derives the control zone name from the fan-coil name by removing " Fan Coil Unit".

Prerequisites (required placeholders)
- The IDF must contain one or more ZoneHVAC:FourPipeFanCoil objects.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class FanCoilNightCycleOperation : ScriptBase, IScript
    {
        // ----------------------------
        // USER CONFIGURATION SECTION
        // ----------------------------
        // IDF template for AvailabilityManagerAssignmentList and AvailabilityManager:NightCycle
        private static readonly string NightCycleIdfTemplate = @"
AvailabilityManagerAssignmentList,
    {0},                                  !- Name
    AvailabilityManager:NightCycle,       !- Availability Manager 1 Object Type
    {1} Night Cycle Operation;            !- Availability Manager 1 Name

AvailabilityManager:NightCycle,
    {1} Night Cycle Operation,            !- Name
    On 24/7,                              !- Applicability Schedule Name
    {3},                                  !- Fan Schedule Name
    CycleOnControlZone,                   !- Control Type
    1,                                    !- Thermostat Tolerance deltaC
    FixedRunTime,                         !- Cycling Run Time Control Type
    3600,                                 !- Cycling Run Time s
    {2};                                  !- Control zone name
";

        private IdfObject FindObject(IdfReader reader, string objectType, string objectName)
        {
            try
            {
                return reader[objectType].First(o => o[0] == objectName);
            }
            catch
            {
                throw new Exception(string.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            StringBuilder idfContent = new StringBuilder();

            IEnumerable<IdfObject> fanCoils = idfReader["ZoneHVAC:FourPipeFanCoil"];
            foreach (IdfObject fanCoil in fanCoils)
            {
                // Add assignment list and nigt cycle manager to the idf
                string fanCoilName = fanCoil["Name"].Value;
                string assignmentListName = fanCoilName + " Assignment List";
                string zoneName = fanCoilName.Replace(" Fan Coil Unit", "");

                // Attach the availability manager list to the fan-coil unit
                fanCoil.AddField(assignmentListName, "!- Availability Manager List Name");

                // Force the unit schedule to always available so Night Cycle can control the fan operation
                fanCoil["Availability Schedule Name"].Value = "On 24/7";

                // Get fan object and its availability schedule (used by AvailabilityManager:NightCycle)
                IdfObject fan = FindObject(
                    idfReader,
                    fanCoil["Supply Air Fan Object Type"].Value,
                    fanCoil["Supply Air Fan Name"].Value);

                string fanScheduleName = fan["Availability Schedule Name"].Value;

                idfContent.AppendFormat(
                    NightCycleIdfTemplate,
                    assignmentListName,
                    fanCoilName,
                    zoneName,
                    fanScheduleName);

                idfContent.AppendLine();
            }

            idfReader.Load(idfContent.ToString());
            idfReader.Save();
        }
    }
}