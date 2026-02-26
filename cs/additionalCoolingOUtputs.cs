/*
 Adds additional EnergyPlus output variables to the cooling
 simulation IDF. Actions performed by the script:
    - Loads an Output:Variable for "People Latent Gain Rate" at Timestep
        reporting frequency.
    - Saves the modified IDF before the cooling simulation runs.
 Usage:
    - Attach this script to run before the cooling simulation.
*/

using System.Runtime;
using System.Collections.Generic;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        public override void BeforeCoolingSimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            idfReader.Load("Output:Variable, *, People Latent Gain Rate, Timestep;");
            idfReader.Save();
        }
    }
}