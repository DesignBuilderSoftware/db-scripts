/*
Include cooling and heating sizing in the DesignBuilder results set.

Purpose:
This DesignBuilder C# script enables "Run Simulation for Sizing Periods" in EnergyPlus so that heating and
cooling sizing-related outputs are available and can be accessed via DesignBuilder Results Viewer.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. 
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class EnableSizingPeriodsInResults : ScriptBase, IScript
    {
        public override void BeforeEnergySimulation()
        {
            IdfReader idf = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // Enable sizing period simulations so sizing-related outputs can appear in results.
            IdfObject simulationControl = idf["SimulationControl"][0];
            simulationControl["Run Simulation for Sizing Periods"].Value = "Yes";

            idf.Save();
        }
    }
}