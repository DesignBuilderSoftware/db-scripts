/*
Night Ventilation Availability Manager Assignment Script

Purpose:
This DesignBuilder C# script adds an EnergyPlus AvailabilityManager:NightVentilation controller to one or more specified air loops.

Main Steps:
1) Locate each target air loop’s supply fan (via the air loop main Branch object)
2) Extract the fan Availability Schedule Name
3) Create an AvailabilityManager:NightVentilation object using that fan schedule and a user-specified control zone
4) Assignthe new AvailabilityManager to the air loop’s AvailabilityManagerAssignmentList
5) Add shared schedules used by the night ventilation manager (Applicability + Setpoint)

How to Use:

Configuration
- Target selection: air loops are selected by the dictionary keys in loopControlZonePairs.
- Control zone: set per air loop using the dictionary values.

Prerequisites / Placeholders
- One or more Air Loops must be included in the base model.
- The air lopps must have supply fans which support the night ventilation control (Fan:VariableVolume, Fan:ConstantVolume, Fan:OnOff)

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. 
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        // Define IDF templates
        //  Parameterized fields: {0} = manager name, {1} = fan availability schedule name, {2} = control zone name
        private readonly string nightVentilationTemplate = @"
AvailabilityManager:NightVentilation,
    {0},                     !- Name
    Night Ventilation Applicability Schedule, !- Applicability Schedule Name
    {1},                     !- Fan Schedule Name
    Night Ventilation Setpoint Schedule,  !- Ventilation Temperature Schedule Name
    2,                       !- Ventilation Temperature Difference deltaC
    15,                      !- Ventilation Temperature Low Limit C
    0.5,                     !- Night Venting Flow Fraction
    {2};                     !- Control Zone Name";

        // Shared schedules used by AvailabilityManager:NightVentilation.
        private readonly string sharedSchedulesIdf = @"
Schedule:Compact,
    Night Ventilation Setpoint Schedule,  !- Name
    Any Number,              !- Schedule Type Limits Name
    Through: 12/31,          !- Field 1
    For: AllDays,            !- Field 2
    Until: 24:00,            !- Field 3
    24;                      !- Field 4

Schedule:Compact,
    Night Ventilation Applicability Schedule,  !- Name
    Any Number,              !- Schedule Type Limits Name
    Through: 12/31,          !- Field 1
    For: AllDays,            !- Field 2
    Until: 24:00,            !- Field 3
    1;                       !- Field 4";

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

        private string FindFanScheduleName(IdfReader reader, string airLoopName)
        {
            IdfObject mainBranch = FindObject(reader, "Branch", airLoopName + " AHU Main Branch");
            // Supported fan types to search for within the Branch equipment list fields
            string[] fanObjectTypes = { "Fan:VariableVolume", "Fan:ConstantVolume", "Fan:OnOff" };
            var fanRefField = mainBranch
                .Select((field, idx) => new { field, idx })
                .FirstOrDefault(x => fanObjectTypes.Contains(x.field.Value));

            if (fanRefField == null)
            {
                throw new Exception("Cannot find fan in air loop: " + airLoopName);
            }

            int fanTypeFieldIndex = fanRefField.idx;
            string fanObjectType = mainBranch[fanTypeFieldIndex].Value;
            string fanObjectName = mainBranch[fanTypeFieldIndex + 1].Value;
            IdfObject fan = FindObject(reader, fanObjectType, fanObjectName);
            return fan[1].Value;
        }

        private void UpdateAssignmentList(IdfReader reader, string airLoopName, string managerObjectType, string managerInstanceName)
        {
            string assignmentListObjectType = "AvailabilityManagerAssignmentList";
            string assignmentListName = airLoopName + " AvailabilityManager List";

            IdfObject assignmentList = FindObject(reader, assignmentListObjectType, assignmentListName);
            if (assignmentList[1] == "AvailabilityManager:Scheduled")
            {
                assignmentList[1].Value = managerObjectType;
                assignmentList[2].Value = managerInstanceName;
            }
            else
            {
                assignmentList.AddFields(managerObjectType, managerInstanceName);
            }
        }

        public override void BeforeEnergySimulation()
        {
            IdfReader idf = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // Map: Define key (air loop name) value (control zone name) pairs (check IDF to get exact names)
            var loopControlZonePairs = new Dictionary<string, string>
            {
                { "Air Loop", "Block1:Zone1" },
                // { "Air Loop 1", "Block1:Zone2" }
            };

            StringBuilder idfToAdd = new StringBuilder();

            foreach (var pair in loopControlZonePairs)
            {
                string airLoopName = pair.Key;
                string controlZoneName = pair.Value;

                string nightVentManagerObjectType = "AvailabilityManager:NightVentilation";
                string nightVentManagerName = airLoopName + " Night Ventilation";

                // Determine the fan availability schedule by locating the supply fan on the air loop main branch
                string fanScheduleName = FindFanScheduleName(idf, airLoopName);

                idfToAdd.AppendFormat(nightVentilationTemplate, nightVentManagerName, fanScheduleName, controlZoneName);
                idfToAdd.Append(Environment.NewLine);

                UpdateAssignmentList(idf, airLoopName, nightVentManagerObjectType, nightVentManagerName);
            }

            idfToAdd.Append(sharedSchedulesIdf);

            idf.Load(idfToAdd.ToString());
            idf.Save();
        }
    }
}