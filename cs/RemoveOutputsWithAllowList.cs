/*
Remove Environment Outputs

Purpose:
This DesignBuilder C# script removes all existing EnergyPlus output requests for Output:Variable and Output:Meter from the IDF,
then adds back only an allow-list of Output:Variable requests so the simulation output (ESO/SQLite/CSV depending on OutputControl:Files) 
contains just the requested variables.

Main Steps:
1) Capture the "Reporting Frequency" values already present in Output:Variable objects (as configured by DesignBuilder).
2) Remove all Output:Variable and Output:Meter* (Meter, MeterFileOnly, Cumulative, Cumulative:MeterFileOnly) objects.
3) Re-add Output:Variable objects for each variable name in the allow-list, for each captured reporting frequency.
4) Replace OutputControl:Files to enforce only the file outputs specified in this script.
5) Save the modified IDF.

How to Use:

Configuration
- allowedOutputVariableNames: list of allowed Output:Variable names (all other Output:Variable requests are removed).
- OutputControl:Files block: controls which output files are produced (CSV/SQLite/ESO/etc).

Prerequisites / Placeholders
- Variable names in allowedOutputVariableNames must match EnergyPlus variable names exactly.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script.
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Linq;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class RemoveEnvironmentOutputs : ScriptBase, IScript
    {
        // ---------------------------
        // USER CONFIGURATION SECTION
        // ---------------------------
        // Allow-list of Output:Variable names to request (all existing Output:Variable objects will be removed first).
        // NOTE: These names must match EnergyPlus Output:Variable names exactly.
        string[] allowedOutputVariableNames = new string[]
        {
            "Site Outdoor Air Drybulb Temperature",
            "Site Outdoor Air Dewpoint Temperature",
            "Site Wind Speed",
            "Site Wind Direction",
            "Site Direct Solar Radiation Rate per Area",
            "Site Diffuse Solar Radiation Rate per Area",
            "Site Solar Azimuth Angle",
            "Site Solar Altitude Angle",
            "Site Outdoor Air Barometric Pressure",
            "Zone Mechanical Ventilation Current Density Volume Flow Rate",
            "Zone Mechanical Ventilation Mass Flow Rate",
            "Air System Outdoor Air Mass Flow Rate",
            "Air System Outdoor Air Flow Fraction",
            "Heat Exchanger Sensible Heating Rate",
            "Heat Exchanger Total Heating Rate",
            "Heat Exchanger Total Cooling Rate",
            "Heat Exchanger Electricity Rate",
            "Zone Air System Sensible Cooling Rate",
            "Zone Air System Sensible Heating Rate"
        };

        public override void BeforeEnergySimulation()
        {
            IdfReader idf = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            var outputVariables = idf["Output:Variable"];

            // Capture reporting frequencies that were already enabled by DesignBuilder in the base IDF.
            var reportingFrequencies = outputVariables
                .Select(outputVar =>
                {
                    var field = outputVar["Reporting Frequency"];
                    return field != null && field.Value != null ? field.Value.Trim() : "";
                })
                .Where(frequency => !string.IsNullOrWhiteSpace(frequency))
                .Distinct()
                .ToList();

            // Remove all existing Output:Variable objects (so we can rebuild them from the allow-list).
            foreach (var outputVar in outputVariables.ToList())
            {
                idf.Remove(outputVar);
            }

            // Remove all meter output request objects (all types used by EnergyPlus for meters).
            foreach (var meterObj in idf["Output:Meter"].ToList())
            {
                idf.Remove(meterObj);
            }

            foreach (var meterObj in idf["Output:Meter:MeterFileOnly"].ToList())
            {
                idf.Remove(meterObj);
            }

            foreach (var meterObj in idf["Output:Meter:Cumulative"].ToList())
            {
                idf.Remove(meterObj);
            }

            foreach (var meterObj in idf["Output:Meter:Cumulative:MeterFileOnly"].ToList())
            {
                idf.Remove(meterObj);
            }

            // Re-add allow-listed Output:Variable objects for each reporting frequency found in the original IDF.
            foreach (var frequency in reportingFrequencies)
            {
                foreach (var variableName in allowedOutputVariableNames)
                {
                    idf.Load(String.Format("Output:Variable,*,{0},{1};", variableName, frequency));
                }
            }

            // Replace any existing OutputControl:Files object to enforce the output file settings below.
            foreach (var filesObj in idf["OutputControl:Files"].ToList())
            {
                idf.Remove(filesObj);
            }

            // ---------------------------
            // USER CONFIGURATION SECTION
            // ---------------------------
            // OutputControl:Files settings (only file types marked as 'Yes' will be saved)
            // NOTE: Keep END and SQLite on for safety / compatibility
            idf.Load(
@"OutputControl:Files,
Yes,  !- CSV
No,   !- MTR
No,   !- ESO
No,   !- EIO
No,   !- Tabular
Yes,  !- SQLite
No,   !- JSON
No,   !- AUDIT
No,   !- Zone Sizing
No,   !- System Sizing
No,   !- DXF
No,   !- BND
No,   !- RDD
No,   !- MDD
No,   !- MTD
Yes,  !- END
No,   !- SHD
No,   !- DFS
No,   !- GLHE
No,   !- DelightIn
No,   !- DelightELdmp
No,   !- DelightDFdmp
No,   !- EDD
No,   !- DBG
No,   !- PerfLog
No,   !- SLN
No,   !- SCI
No,   !- WRL
No,   !- Screen
No,   !- ExtShd
No;   !- Tarcog");

            idf.Save();
        }
    }
}