/* 
Override Daylighting controls schedule to disable daylighting control for sizing calculations.

*/
using System.Collections.Generic;
using System.Windows.Forms;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        private void UpdateDlSchedule()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            IEnumerable<IdfObject> dlControls = idfReader["Daylighting:Controls"];

            MessageBox.Show("Updating daylighting control availability");

            foreach (IdfObject dlControl in dlControls)
            {
                dlControl["Availability Schedule Name"].Value = "OnSddOff";
            }

            idfReader.Load("Schedule:Compact,\nOnSddOff,\nFraction,\nThrough: 12/31,\nFor: SummerDesignDay,\nUntil: 24:00,\n0,\nFor: AllOtherDays,\nUntil: 24:00,\n1;");

            idfReader.Save();
        }

        public override void BeforeEnergySimulation()
        {
            UpdateDlSchedule();
        }

        public override void BeforeCoolingSimulation()
        {
            UpdateDlSchedule();
        }
    }

}