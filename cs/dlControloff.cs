/*
Daylighting Controls Sizing Override (Disable Daylighting on Cooling Design Days)

Purpose
This DesignBuilder C# script disables daylighting controls for sizing-related calculations by
overriding the Availability Schedule referenced by all Daylighting:Controls objects.

1) Find all Daylighting:Controls objects
2) Force their "Availability Schedule Name" to a custom schedule (default: "OnSddOff")
3) Inject a Schedule:Compact object named "OnSddOff" that:
  - Sets availability to 0 during SummerDesignDay (daylighting disabled for sizing design day)
  - Sets availability to 1 for all other days (daylighting enabled for normal simulation periods)
4) Save the modified IDF back to disk before EnergyPlus runs

How to Use
- Add this script to your DesignBuilder model as a C# extensibility script.
- The script runs in these hooks:
  - BeforeEnergySimulation()
  - BeforeCoolingSimulation()
  This ensures the IDF is updated before the relevant simulation phases that may include sizing.

Configuration
- Schedule name: "OnSddOff"

Prerequisites
- The model/IDF must contain one or more Daylighting:Controls objects.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System.Collections.Generic;
using System.Windows.Forms;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        private void UpdateDlSchedule()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // Target all daylighting control objects and override their availability schedule reference.
            IEnumerable<IdfObject> daylightingControls = idfReader["Daylighting:Controls"];

            MessageBox.Show("Updating daylighting control availability"); // Comment this line to disable the message box.

            // Force each daylighting control object to use the override schedule.
            foreach (IdfObject daylightingControl in daylightingControls)
            {
                daylightingControl["Availability Schedule Name"].Value = "OnSddOff";
            }

            // "OnSddOff" is defined as: 0 on SummerDesignDay, 1 on all other days.
            const string onSddOffScheduleCompact =
                "Schedule:Compact,\n" +
                "OnSddOff,\n" +
                "Fraction,\n" +
                "Through: 12/31,\n" +
                "For: SummerDesignDay,\n" +
                "Until: 24:00,\n" +
                "0,\n" +
                "For: AllOtherDays,\n" +
                "Until: 24:00,\n" +
                "1;";

            idfReader.Load(onSddOffScheduleCompact);

            idfReader.Save();
        }

        public override void BeforeEnergySimulation()
        {
            UpdateDlSchedule();
        }

        public override void BeforeCoolingSimulation()
        {
            UpdateDlSchedule();
        }
    }
}