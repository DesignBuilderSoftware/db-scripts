/*
Infiltration Conversion Script (Flow/ExteriorWallArea)

This DesignBuilder C# script converts ZoneInfiltration:DesignFlowRate objects to use the "Flow/ExteriorWallArea" calculation method
and populates the "Flow per Exterior Surface Area" field using the infiltration value defined in the DesignBuilder UI.

Purpose (main steps)
1) Retrieve each zone’s infiltration value from DesignBuilder attributes (InfiltrationValueI4).
2) Update every ZoneInfiltration:DesignFlowRate object:
   - Set Design Flow Rate Calculation Method = "Flow/ExteriorWallArea"
   - Set Flow per Exterior Surface Area = (DB value) / 3600  [m3/h-m2 to m3/s-m2]
3) Save the modified IDF back to disk before EnergyPlus starts.

How to Use

Configuration
- The infiltration rate is retrieved from DesignBuilder interface.
    This setting can be controled in 'Construction > Airtightness > Model Infiltration'.

Prerequisites (required placeholders)
- This script requires the DesignBuilder "Infiltration units" option set to "3 - m3/h-m2 at 4Pa".
    This setting can be controled in 'Model Options > Natural Ventilation and Infiltration > Infiltration Units'.

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
    public class InfiltrationChange : ScriptBase, IScript
    {
        public override void BeforeEnergySimulation()
        {
            IdfReader idf = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath
            );
            Site site = ApiEnvironment.Site;
            Building building = site.Buildings[ApiEnvironment.CurrentBuildingIndex];

            // Read infiltration values from DesignBuilder zone attributes (InfiltrationValueI4)
            Dictionary<string, string> zoneInfiltrationValueI4ByZoneName =
                GetZoneInfiltrationValueI4ByZoneName(building);

            // Target EnergyPlus objects: ZoneInfiltration:DesignFlowRate
            IEnumerable<IdfObject> zoneInfiltrationDesignFlowObjects =
                idf["ZONEINFILTRATION:DESIGNFLOWRATE"];

            foreach (IdfObject zoneInfilObj in zoneInfiltrationDesignFlowObjects)
            {
                string zoneName = zoneInfilObj[1];

                // Force EnergyPlus to use exterior wall area-based infiltration.
                zoneInfilObj["Design Flow Rate Calculation Method"].Value = "Flow/ExteriorWallArea";

                // Convert from m3/h-m2 (DesignBuilder UI) to m3/sm2 (EnergyPlus field units)
                zoneInfilObj["Flow per Exterior Surface Area"].Value =
                    (double.Parse(zoneInfiltrationValueI4ByZoneName[zoneName]) / 3600.0).ToString();
            }

            idf.Save();
        }

        private Dictionary<string, string> GetZoneInfiltrationValueI4ByZoneName(Building building)
        {
            return building.BuildingBlocks
                .SelectMany(block => block.Zones)
                .ToDictionary(
                    zone => zone.IdfName,
                    zone => zone.GetAttribute("InfiltrationValueI4")
                );
        }
    }
}