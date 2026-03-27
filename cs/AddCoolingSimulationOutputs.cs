/*
Adds additional EnergyPlus output variables to the cooling simulation IDF

Purpose:
This DesignBuilder C# script adds an EnergyPlus Output:Variable request to the cooling simulation IDF.

Main Steps:
1) Insert an Output:Variable object requesting variables at a given frequency
2) Save the modified IDF so EnergyPlus uses the added output request during the cooling simulation

How to Use:

Configuration
- Add an Output:Variable definition in this order:
    - Object Name: Output:Variable
    - Key Value: The specific instance name (e.g., Zone Name, Component Name). Use "*" to request the variable for all applicable keys.
    - Variable Name: The exact name of the EnergyPlus output requested, as found in the .rdd file (e.g., "People Latent Gain Rate").
    - Reporting Frequency: the desired reporting interval supported by EnergyPlus ("Timestep", "Hourly", "Daily", "Monthly", "RunPeriod", "Annual")

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. 
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class AddCoolingPeopleLatentGainOutput : ScriptBase, IScript
    {
        public override void BeforeCoolingSimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // ----------------------------
            // USER CONFIGURATION SECTION
            // ----------------------------
            // Add Output:Variable requests
            idfReader.Load("Output:Variable, *, People Latent Gain Rate, Timestep;");

            idfReader.Save();
        }
    }
}