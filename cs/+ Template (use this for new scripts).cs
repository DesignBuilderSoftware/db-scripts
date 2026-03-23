/*
<Simple description of the script (one line)>

Purpose:
- A few sentences describing what the script does in DesignBuilder/EnergyPlus in more detail.

Main Steps:
1) <Step 1: what it inspects/reads>
2) <Step 2: what it builds/changes>
3) <Step 3: how it writes changes (IDF / DB model)>

How to Use:

Configuration
- <Explain key configuration variables and typical values>
- <Sections in the script should be marked with a "USER CONFIGURATION SECTION" flag>

Prerequisites / Placeholders
- <Required placeholders such as DB objects, attributes and/or settings that should be adjusted via its interface>

Notes:
- <General notes that might be useful to the user> 
- <Important assumptions, edge cases, what it will NOT do>

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. 
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

// < Required namespaces>
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Globalization;
using DB.Extensibility.Contracts;
using DB.Api;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class MyCustomScriptTemplate : ScriptBase, IScript
    {
        // ---------------------------
        // USER CONFIGURATION SECTION <Indicate with this flag the sections which the user should check and change>
        // ---------------------------
        // Define EnergyPlus object templates as strings (Boilerplates)
        private readonly string energyPlusObjectBoilerplate = @"
            Component:Type,
            {0},                     !- Name
            {1},                     !- Field 1
            {2};                     !- Field 2
        ";

        // Hook Point: Executes before the EnergyPlus simulation starts
        public override void BeforeEnergySimulation()
        {
            // 1. Initialize the IdfReader to modify the simulation input
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // 2. Access the DesignBuilder API to retrieve model data
            Site site = ApiEnvironment.Site;
            Building building = site.Buildings[ApiEnvironment.CurrentBuildingIndex];

            // 3. Script Logic: Manipulation of IDF Objects [15], [16], [17]
            // Example: Iterate through existing objects or add new ones
            foreach (var idfObject in idfReader["AirLoopHVAC"].ToList())
            {
                // Perform modifications or data extraction
            }

            // 4. Load new boilerplate objects into the reader
            // idfReader.Load(String.Format(energyPlusObjectBoilerplate, "Name", "Value1", "Value2"));

            // 5. Save the modified IDF for the simulation engine to use
            idfReader.Save();
        }
    }
}