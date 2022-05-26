/*
Replace electric steam humidifiers with equivalent gas component.

Air loop needs to contain 'steamgas' keyword in order to be updated.

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

        private void UpdateBranches(IdfReader idfReader)
        {
             IEnumerable<IdfObject> branches = idfReader["Branch"];
             foreach (IdfObject branch in branches)
             {
                 if (branch[0].Value.Contains("steamgas") && branch[0].Value.Contains("AHU Main Branch"))
                 {
                      UpdateBranch(idfReader, branch);
                 }
             }
        }

        private void UpdateBranch(IdfReader reader, IdfObject branch)
        {
             const string GasHumidifierObject = "Humidifier:Steam:Gas";
             const string EleHumidifierObject = "Humidifier:Steam:Electric";

             foreach (int i in Enumerable.Range(1, (branch.Count - 2)))
             {
                   Field ThisField = branch.Fields[i];
                   Field NextField = branch.Fields[i + 1];
                   if (ThisField.Comment.ToLower().Contains("object type") && ThisField.Value.ToLower() == EleHumidifierObject.ToLower())
                   {
                        IdfObject EleHumidifier = FindObject(reader, EleHumidifierObject, NextField.Value);
                        ThisField.Value = GasHumidifierObject;
                        string humidifierText = GetHumidifierIdfText(EleHumidifier, GasHumidifierObject);
                        reader.Load(humidifierText);
                        MessageBox.Show("Replacing electric humidifier: " + EleHumidifier[0] + " with gas humidifier." );
                        reader.Remove(EleHumidifier);
                        break;
                   }
             }
        }
        private string GetHumidifierIdfText(IdfObject EleHumidifier, string GasHumidifier)
        {
            string name = EleHumidifier[0].Value;
            string availability = EleHumidifier[1].Value;
            string ratedCapacity = EleHumidifier[2].Value;
            string ratedGasRate = EleHumidifier[3].Value;
            string thermalEfficiency = "0.9";
            string thermalEfficiencyCurve = "";
            string ratedFanPower = EleHumidifier[4].Value;
            string auxPower = EleHumidifier[5].Value;
            string inlet = EleHumidifier[6].Value;
            string outlet = EleHumidifier[7].Value;

            string[] fields = { GasHumidifier, name, availability, ratedCapacity, ratedGasRate, thermalEfficiency, thermalEfficiencyCurve, ratedFanPower, auxPower, inlet, outlet };
            return String.Join(",", fields) + ";";
        }

        private void AddOutputs(IdfReader reader)
        {
            reader.Load("Output:Variable,*,Humidifier NaturalGas Rate,hourly;");
            reader.Load("Output:Variable,*,Humidifier NaturalGas Rate,Runperiod;");
        }

         public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            UpdateBranches(idfReader);
            AddOutputs(idfReader);

            idfReader.Save();
        }
    }
}
