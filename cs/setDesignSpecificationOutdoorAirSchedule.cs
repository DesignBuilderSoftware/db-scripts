/*
UpdateDesignSpecification – Apply Mechanical Ventilation “Operation Schedule” to Detailed HVAC Outdoor Air Specs

Purpose
This DesignBuilder C# script reads each zone’s Mechanical Ventilation “Operation Schedule” from the model
and assigns the corresponding EnergyPlus schedule name to the zone’s DesignSpecification:OutdoorAir object.

Main Steps
1) Read the DesignBuilder “Schedules” table (model data).
2) Build a map of Zone IDF Name to MechanicalVentilationSchedule handle (string) from model attributes.
3) Parse the EnergyPlus IDF to find existing schedule names (Schedule:Compact).
4) For each DesignSpecification:OutdoorAir:
   - Derive the zone name from the object name (assumes standard naming convention).
   - Find the schedule record in the “Schedules” table using the stored handle.
   - If the schedule does not exist in the IDF, load it from the schedule CompactData.
   - Set “Outdoor Air Schedule Name” to that schedule name.
NOTE: he script currently only checks Schedule:Compact objects.

How to Use

Prerequisites (required placeholders)
- Zones must have the attribute “MechanicalVentilationSchedule” populated via model user interface.
- The EnergyPlus IDF must contain DesignSpecification:OutdoorAir objects created by Detailed HVAC.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Collections.Generic;
using System.Linq;

using DB.Api;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class UpdateDesignSpecification : ScriptBase, IScript
    {
        public override void BeforeEnergySimulation()
        {
            Site site = ApiEnvironment.Site;
            Table schedulesTable = site.GetTable("Schedules");

            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // Build mapping: Zone IdfName -> schedule record handle stored in the zone attribute
            Building building = site.Buildings[ApiEnvironment.CurrentBuildingIndex];
            Dictionary<string, string> zoneToScheduleHandle = GetZoneOAScheduleIds(building);
            HashSet<string> scheduleNames = GetScheduleNames(idfReader);

            string result = string.Join(", ", zoneToScheduleHandle.Select(x => x.Key + ": " + x.Value));

            // All DesignSpecification:OutdoorAir objects that will receive “Outdoor Air Schedule Name”
            IEnumerable<IdfObject> outdoorAirDesignSpecs = idfReader["DesignSpecification:OutdoorAir"];

            foreach (IdfObject designSpecification in outdoorAirDesignSpecs)
            {
                string zoneName = designSpecification["Name"].Value.Replace(" Design Specification Outdoor Air Object", "");
                string scheduleHandle = zoneToScheduleHandle[zoneName];

                // Retrieve schedule record from model data using its handle
                Record scheduleRecord = schedulesTable.Records.GetRecordFromHandle(int.Parse(scheduleHandle));

                // Schedule name as it should appear in the EnergyPlus IDF
                string scheduleName = scheduleRecord["Name"];

                // If schedule is not already in the IDF, reconstruct Schedule:Compact text and load it
                if (!scheduleNames.Contains(scheduleName))
                {
                    string scheduleContent = ParseScheduleData(scheduleRecord["CompactData"], scheduleName);
                    idfReader.Load(scheduleContent);
                    scheduleNames.Add(scheduleName);
                }

                // Apply the schedule to the DesignSpecification:OutdoorAir object
                designSpecification["Outdoor Air Schedule Name"].Value = scheduleName;
            }

            idfReader.Save();
        }

        private string ParseScheduleData(string content, string name)
        {
            // CompactData appears to encode newlines using the two-character delimiter.
            // Replace that delimiter with actual line breaks so it becomes valid IDF text.
            char pipe = (char)124;   // ASCII code for vertical bar: |
            char caret = (char)94;   //  ASCII code for caret: ^
            string delimiter = new string(new[] { pipe, caret });
            content = content.Replace(delimiter, System.Environment.NewLine);

            // Replace the schedule name in the Schedule:Compact object.
            int firstCommaIndex = content.IndexOf(',');
            int secondCommaIndex = content.IndexOf(',', firstCommaIndex + 1);

            string beforeFirstComma = content.Substring(0, firstCommaIndex + 1);
            string afterSecondComma = content.Substring(secondCommaIndex);

            return beforeFirstComma + name + " " + afterSecondComma;
        }

        private HashSet<string> GetScheduleNames(IdfReader reader)
        {
            string[] scheduleTypes = new string[] { "Schedule:Compact" };

            HashSet<string> scheduleNames = new HashSet<string>(
                scheduleTypes
                    .SelectMany(scheduleType => reader[scheduleType]
                        .Select(schedule => schedule[0].Value)));

            return scheduleNames;
        }

        private Dictionary<string, string> GetZoneOAScheduleIds(Building building)
        {
            // Read the per-zone attribute “MechanicalVentilationSchedule” which stores a handle (string)
            Dictionary<string, string> zoneOAScheduleIds = building.BuildingBlocks
                .SelectMany(block => block.Zones)
                .ToDictionary(
                    zone => zone.IdfName,
                    zone => zone.GetAttribute("MechanicalVentilationSchedule")
                );

            return zoneOAScheduleIds;
        }
    }
}