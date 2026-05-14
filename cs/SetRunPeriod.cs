/*
Updates the EnergyPlus RunPeriod object to a user-defined date range and weather-file settings.
Version: 1.0

Purpose:
- This script replaces the RunPeriod object in the exported IDF so that the simulation runs over
  a custom start/end date range instead of the default DesignBuilder-generated period.
- It also controls which weather-file flags (holidays, daylight saving, rain, snow) are applied.

Main Steps:
1) Locate the existing RunPeriod object in the IDF and read its name.
2) Build a replacement RunPeriod object string from the configured date and flag inputs.
3) Load the new RunPeriod into the IDF, remove the old one to prevent duplicates, and save.

How to Use:

Configuration
- beginMonth / beginDay / beginYear: simulation start date (integer month/day/year).
- endMonth / endDay / endYear: simulation end date (integer month/day/year).
- dayOfWeek: EnergyPlus day-of-week string for the start day
    Valid values: "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday",
    "SummerSolstice", "WinterSolstice", "UseWeatherFile".
- useWeatherFileHolidays: whether to import holidays from the weather file ("Yes" / "No").
- useWeatherFileDst: whether to apply daylight saving from the weather file ("Yes" / "No").
- applyWeekendHolidayRule: shift Saturday/Sunday holidays to Monday ("Yes" / "No").
- useWeatherFileRain: apply rain indicators from the weather file ("Yes" / "No").
- useWeatherFileSnow: apply snow indicators from the weather file ("Yes" / "No").

Prerequisites / Placeholders
- The simulation should run with "Simulation Manager" enabled.
IMPORTANT: If multiple years are simulated, results will be displayed in Results Viewer only (not in DB interface).

DISCLAIMER: This script is provided as-is without warranty. 
DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script.
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Linq;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class SetRunPeriod : ScriptBase, IScript
    {
        // Hook Point: Executes before the EnergyPlus simulation starts
        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // ----------------------------
            // USER CONFIGURATION SECTION
            // ----------------------------
            UpdateRunPeriod(
                idfReader,
                beginMonth: 9,
                beginDay: 18,
                beginYear: 2018,
                endMonth: 9,
                endDay: 17,
                endYear: 2019,
                dayOfWeek: "Tuesday",
                useWeatherFileHolidays: "No",
                useWeatherFileDst: "Yes",
                applyWeekendHolidayRule: "No",
                useWeatherFileRain: "Yes",
                useWeatherFileSnow: "Yes");

            idfReader.Save();
        }

        // Replaces the existing RunPeriod object with a new one built from the supplied inputs.
        private void UpdateRunPeriod(
            IdfReader idfReader,
            int beginMonth, int beginDay, int beginYear,
            int endMonth, int endDay, int endYear,
            string dayOfWeek,
            string useWeatherFileHolidays,
            string useWeatherFileDst,
            string applyWeekendHolidayRule,
            string useWeatherFileRain,
            string useWeatherFileSnow)
        {
            // Read the existing RunPeriod so its name can be reused in the replacement object
            IdfObject existingRunPeriod = idfReader["RunPeriod"].First();
            string runPeriodName = existingRunPeriod[0];

            string newRunPeriodIdfText = BuildRunPeriodIdfText(
                runPeriodName,
                beginMonth, beginDay, beginYear,
                endMonth, endDay, endYear,
                dayOfWeek,
                useWeatherFileHolidays, useWeatherFileDst, applyWeekendHolidayRule,
                useWeatherFileRain, useWeatherFileSnow);

            // Load the replacement first, then remove the old one to prevent duplicate-object errors
            idfReader.Load(newRunPeriodIdfText);
            idfReader.Remove(existingRunPeriod);
        }

        // Returns the IDF text for a RunPeriod object built from the provided configuration values.
        private string BuildRunPeriodIdfText(
            string runPeriodName,
            int beginMonth, int beginDay, int beginYear,
            int endMonth, int endDay, int endYear,
            string dayOfWeek,
            string useWeatherFileHolidays,
            string useWeatherFileDst,
            string applyWeekendHolidayRule,
            string useWeatherFileRain,
            string useWeatherFileSnow)
        {
            string runPeriodTemplate = @"RunPeriod,
  {0},   !- Name
  {1},   !- Begin Month
  {2},   !- Begin Day of Month
  {3},   !- Begin Year
  {4},   !- End Month
  {5},   !- End Day of Month
  {6},   !- End Year
  {7},   !- Day of Week for Start Day
  {8},   !- Use Weather File Holidays and Special Days
  {9},   !- Use Weather File Daylight Saving Period
  {10},  !- Apply Weekend Holiday Rule
  {11},  !- Use Weather File Rain Indicators
  {12};  !- Use Weather File Snow Indicators";

            return String.Format(
                runPeriodTemplate,
                runPeriodName,
                beginMonth, beginDay, beginYear,
                endMonth, endDay, endYear,
                dayOfWeek,
                useWeatherFileHolidays, useWeatherFileDst, applyWeekendHolidayRule,
                useWeatherFileRain, useWeatherFileSnow);
        }
    }
}
