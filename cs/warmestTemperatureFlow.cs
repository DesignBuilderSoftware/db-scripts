/*
Replace SetpointManager:Warmest with SetpointManager:WarmestTemperatureFlow.

Setpoint manager name must contain 'WarmestTemperatureFlow' keyword in order to be replaced.

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

        private IdfObject FindObject(IdfReader reader, string objectType, string objectName) {
             return reader[objectType].First(c => c[0] == objectName);
        }

        private void ReplaceWarmest(IdfReader idfReader, string strategy, float turnDownRatio)
        {
             IEnumerable<IdfObject> spms = idfReader["SetpointManager:Warmest"];
             foreach (IdfObject spm in spms)
             {
                 if (spm[0].Value.ToLower().Contains("warmesttemperatureflow"))
                 {
                      string newSpm = GetSpmText(spm, strategy, turnDownRatio);
                      MessageBox.Show("Replacing Warmest spm: " + spm[0].Value + " with WarmestTemperatureFlow spm." );
                      idfReader.Load(newSpm);
                      idfReader.Remove(spm);
                 }
             }
        }

        private string GetSpmText(IdfObject warmestSpm, string strategy, float turnDownRatio)
        {
            string objectName = "SetpointManager:WarmestTemperatureFlow";
            string name = warmestSpm[0].Value;
            string controlVariable = warmestSpm[1].Value;
            string airLoop = warmestSpm[2].Value;
            string minTemp = warmestSpm[3].Value;
            string maxTemp = warmestSpm[4].Value;
            string node = warmestSpm[6].Value;
            string[] fields = { objectName, name, controlVariable, airLoop, minTemp, maxTemp, strategy, node, turnDownRatio.ToString() };
            return String.Join(",", fields) + ";";
        }

         public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);
            string strategy = "TemperatureFirst";
            // string strategy = "FlowFirst";
            float turnDownRatio = 0.3f;
            ReplaceWarmest(idfReader, strategy, turnDownRatio);
            idfReader.Save();
        }
    }
}
