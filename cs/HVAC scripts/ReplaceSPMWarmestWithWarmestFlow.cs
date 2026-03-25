/*
Replaces SetpointManager from type Warmest to TemperatureFlow.

Purpose:
This DesignBuilder C# script replaces SetpointManager:Warmest with SetpointManager:WarmestTemperatureFlow.
This object  provides a setpoint manager that attempts to establish a supply air setpoint that will meet 
the cooling load of the zone needing the coldest air at the maximum zone supply air flowrate. 

Main Steps:
1) Find all SetpointManager:Warmest objects
2) Select only those whose Name contains keyword (default: "WarmestTemperatureFlow")
3) Replace each selected SetpointManager:Warmest with a SetpointManager:WarmestTemperatureFlow object
  using the same core fields (Name, Control Variable, Air Loop, Min/Max setpoint, Setpoint Node/NodeList),
  and applying the configured Strategy and Minimum Turndown Ratio.

How to Use:

Configuration
- Choose strategy type: "TemperatureFirst" (default) or "FlowFirst"
- defined Minimum turndown ratio: 0-1

Prerequisites / Placeholders
Base model must already contain at least one SetpointManager:Warmest with name containing "WarmestTemperatureFlow"

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. 
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        private IdfObject FindObject(IdfReader reader, string objectType, string objectName)
        {
            return reader[objectType].First(c => c[0] == objectName);
        }

        private void ReplaceWarmest(IdfReader idfReader, string strategy, float turnDownRatio)
        {
            IEnumerable<IdfObject> warmestSpms = idfReader["SetpointManager:Warmest"];

            foreach (IdfObject warmestSpm in warmestSpms)
            {
                // Selection rule: only replace Warmest SPMs with the keyword in the Name field (default: "WarmestTemperatureFlow")
                if (warmestSpm[0].Value.ToLower().Contains("warmesttemperatureflow"))
                {
                    // Build replacement object text and load into IDF, then remove old object
                    string newSpm = GetSpmText(warmestSpm, strategy, turnDownRatio);

                    MessageBox.Show("Replacing Warmest spm: " + warmestSpm[0].Value + " with WarmestTemperatureFlow spm."); // Comment this line to remove message box

                    idfReader.Load(newSpm);
                    idfReader.Remove(warmestSpm);
                }
            }
        }

        private string GetSpmText(IdfObject warmestSpm, string strategy, float turnDownRatio)
        {
            // Builds a SetpointManager:WarmestTemperatureFlow object line.
            string objectName = "SetpointManager:WarmestTemperatureFlow";
            string name = warmestSpm[0].Value;
            string controlVariable = warmestSpm[1].Value;
            string airLoop = warmestSpm[2].Value;
            string minTemp = warmestSpm[3].Value;
            string maxTemp = warmestSpm[4].Value;
            string node = warmestSpm[6].Value;

            string[] fields =
            {
                objectName,
                name,
                controlVariable,
                airLoop,
                minTemp,
                maxTemp,
                strategy,
                node,
                turnDownRatio.ToString()
            };

            return String.Join(",", fields) + ";";
        }

        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // USER CONFIGURATION: Strategy options ("TemperatureFirst" / "FlowFirst")
            string strategy = "TemperatureFirst";
            // string strategy = "FlowFirst";

            // USER CONFIGURATION: minimum Turndown Ratio (0-1)
            float minimumTurndownRatio = 0.3f;

            ReplaceWarmest(idfReader, strategy, minimumTurndownRatio);
            idfReader.Save();
        }
    }
}