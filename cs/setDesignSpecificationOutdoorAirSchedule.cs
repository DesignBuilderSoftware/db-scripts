/*
Apply the mechanical ventilation "Operation Schedule" in Detailed HVAC.

The script reads the schedule value from model data and assigns it to the "DesignSpecification:OutdoorAir" object.
   
 */
using System;
using System.Runtime;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using EpNet;
using DB.Api;
using DB.Extensibility.Contracts;

namespace DB.Extensibility.Scripts
{
    public class UpdateDesignSpecification : ScriptBase, IScript
    {
        public override void BeforeEnergySimulation()
        {
            Site site = ApiEnvironment.Site;
            Table table = site.GetTable("Schedules");

            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            Building building = site.Buildings[ApiEnvironment.CurrentBuildingIndex];
            Dictionary<string, string> zoneOAScheduleIds = GetZoneOAScheduleIds(building);
            HashSet<string> scheduleNames = GetScheduleNames(idfReader);

            string result = string.Join(", ", zoneOAScheduleIds.Select(x => x.Key + ": " + x.Value));

            IEnumerable<IdfObject> oaDesignSpecs = idfReader["DesignSpecification:OutdoorAir"];

            foreach (IdfObject designSpecification in oaDesignSpecs)
            {
                string zoneName = designSpecification["Name"].Value.Replace(" Design Specification Outdoor Air Object", "");
                string scheduleHandle = zoneOAScheduleIds[zoneName];

                Record scheduleRecord = table.Records.GetRecordFromHandle(int.Parse(scheduleHandle));

                string scheduleName = scheduleRecord["Name"];

                if (!scheduleNames.Contains(scheduleName))
                {
                    string scheduleContent = ParseScheduleData(scheduleRecord["CompactData"], scheduleName);
                    idfReader.Load(scheduleContent);
                    scheduleNames.Add(scheduleName);
                }
                designSpecification["Outdoor Air Schedule Name"].Value = scheduleName;
            }
            idfReader.Save();
        }

        private string ParseScheduleData(string content, string name)
        {
            char pipe = (char)124;  // ASCII code for pipe
            char caret = (char)94;  // ASCII code for caret
            string delimiter = new string(new[] { pipe, caret });
            content = content.Replace(delimiter, System.Environment.NewLine);

            int firstCommaIndex = content.IndexOf(',');
            int secondCommaIndex = content.IndexOf(',', firstCommaIndex + 1);

            string beforeFirstComma = content.Substring(0, firstCommaIndex + 1);
            string afterSecondComma = content.Substring(secondCommaIndex);

            return beforeFirstComma + name + " " + afterSecondComma;
        }

        private HashSet<string> GetScheduleNames(IdfReader reader)
        {
            string[] scheduleTypes = new string[] { "Schedule:Compact" };
            HashSet<string> scheduleNames = new HashSet<string>(scheduleTypes
                .SelectMany(scheduleType => reader[scheduleType]
                .Select(schedule => schedule[0].Value)));
            return scheduleNames;
        }

        private Dictionary<string, string> GetZoneOAScheduleIds(Building building)
        {
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
