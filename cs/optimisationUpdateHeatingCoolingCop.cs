/*
 Reads optimisation variables from the "OptimisationVariables" table
 (expected keys: "heatingCOP" and "coolingEER") and updates the EnergyPlus
 IDF accordingly. Actions performed by the script:
    - Sets the "Rated COP" field on the
        Coil:WaterHeating:AirToWaterHeatPump:Pumped object named
        "HP Water Heater HP Water Heating Coil" when a valid heating COP is present.
    - Sets the "Reference COP" field on Chiller:Electric:EIR objects when a valid cooling EER is present.
 Behavior:
    - If a variable value equals "UNKNOWN", a MessageBox is shown and that value is not applied.
    - The IDF is saved after modifications.
 Usage:
    - Ensure the `OptimisationVariables` table contains the current values before running this script.
*/

using System.Windows.Forms;
using System.Collections.Generic;
using DB.Extensibility.Contracts;
using EpNet;
using DB.Api;
using System.Windows.Forms;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        public override void BeforeEnergySimulation()
        {
            Site site = ApiEnvironment.Site;
            Table table = site.GetTable("OptimisationVariables");
            Record recordHeating = table.Records["heatingCOP"];
            string heatingCop = recordHeating["VariableCurrentValue"];

            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            IEnumerable<IdfObject> heatingCoils = idfReader["Coil:WaterHeating:AirToWaterHeatPump:Pumped"];
            foreach (IdfObject coil in heatingCoils)
            {
                if (coil["Name"].Equals("HP Water Heater HP Water Heating Coil"))
                {
                    if (heatingCop.Equals("UNKNOWN"))
                    {
                        MessageBox.Show("Cannot set heating COP, UNKNOWN value in OptimisationVariables table. ");
                    }
                    else
                    {
                        coil["Rated COP"].Value = heatingCop;
                    }
                }
            }

            Record recordCooling = table.Records["coolingEER"];
            string coolingEer = recordCooling["VariableCurrentValue"];

            IEnumerable<IdfObject> chillers = idfReader["Chiller:Electric:EIR"];
            foreach (IdfObject chiller in chillers)
            {
                if (coolingEer.Equals("UNKNOWN"))
                {
                    MessageBox.Show("Cannot set cooling COP, UNKNOWN value in OptimisationVariables table. ");
                }
                else
                {
                    chiller["Reference COP"].Value = coolingEer;
                }
            }
            idfReader.Save();
        }
    }
}
