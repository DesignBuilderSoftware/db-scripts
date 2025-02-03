/*
Apply the Flow/ExteriorWallArea infiltration method.

The infiltration rate is retrieved from DesignBuilder interface.

This script requires a specific "Infiltration units" option (3-m3/h-m2 at 4Pa).

*/
using System;
using System.Runtime;
using System.Collections.Generic;
using System.Linq;

using DB.Extensibility.Contracts;
using DB.Api;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class InfiltrationChange : ScriptBase, IScript
    {
        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath
            );
            Site site = ApiEnvironment.Site;
            Building building = site.Buildings[ApiEnvironment.CurrentBuildingIndex];
            Dictionary<string, string> zoneInfiltrationRates = GetZoneInfiltrationRates(building);

            IEnumerable<IdfObject> infiltrationObjects = idfReader["ZONEINFILTRATION:DESIGNFLOWRATE"];

            foreach (IdfObject infiltration in infiltrationObjects)
            {
                string zoneName = infiltration[1];
                infiltration["Design Flow Rate Calculation Method"].Value = "Flow/ExteriorWallArea";
                infiltration["Flow Rate per Exterior Surface Area"].Value = zoneInfiltrationRates[zoneName];
            }

            idfReader.Save();
        }

        private Dictionary<string, string> GetZoneInfiltrationRates(Building building)
        {
            Dictionary<string, string> zoneInfiltrationRates = building.BuildingBlocks
                .SelectMany(block => block.Zones)
                .ToDictionary(
                    zone => zone.IdfName,
                    zone => zone.GetAttribute("InfiltrationValueI4")
                );
            return zoneInfiltrationRates;
        }
    }
}