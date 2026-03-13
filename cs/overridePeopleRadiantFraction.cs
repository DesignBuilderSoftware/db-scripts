/*
Override radiant heat gain fraction for People objects.

This DesignBuilder C# script overrides the EnergyPlus People object "Fraction Radiant" field.
DesignBuilder typically applies a default radiant fraction (commonly 0.3). 
This script enforces a user-defined value for all People objects in the IDF just before the EnergyPlus simulation runs.

Purpose
1) Find all People objects
2) Set "Fraction Radiant" to a configured value
3) Save the modified IDF

How to Use

Configuration
- radiantFraction: Radiant fraction of sensible heat gains from people (dimensionless, 0–1).
  This value will be applied to every People object found in the IDF.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System.Collections.Generic;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class OverridePeopleRadiantFraction : ScriptBase, IScript
    {
        public override void BeforeEnergySimulation()
        {
            IdfReader idf = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // ----------------------------
            // USER CONFIGURATION SECTION
            // ----------------------------
            // Target radiant fraction (dimensionless, 0.0 to 1.0).
            const double radiantFraction = 0.15;

            // Retrieve all People objects and set the "Fraction Radiant" field.
            IEnumerable<IdfObject> peopleObjects = idf["People"];
            foreach (IdfObject peopleObj in peopleObjects)
            {
                peopleObj["Fraction Radiant"].Number = radiantFraction;
            }

            idf.Save();
        }
    }
}