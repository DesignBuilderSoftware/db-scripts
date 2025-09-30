/*
Apply desuperheater coil to Unitary Generic AHU systems.

The update is applied to all Unitary Generic AHU systems with desuperheater substring in their name.
The unit must include the CoolReheat humidity control. 
*/

using System.Collections.Generic;
using System.Linq;
using System;
using System.Windows.Forms;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        private IdfReader Reader;

        public IdfObject FindObject(string objectType, string objectName)
        {
            try
            {
                return Reader[objectType].First(c => c[0] == objectName);
            }
            catch (Exception e)
            {
                throw new MissingFieldException(String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        private string GetDesuperheaterCoil(string name, string coolingCoilType, string coolingCoilName, IdfObject reheatCoil)
        {
            string template = @"Coil:Heating:Desuperheater,
  {0},              !- Coil Name
  {1},              !- Availability Schedule
  0.3,              !- Heat Reclaim Recovery Efficiency
  {2},              !- Coil Air Inlet Node Name
  {3},              !- Coil Air Outlet Node Name
  {4},              !- Heating Source Type
  {5},              !- Heating Source Name
  {6},              !- Coil Temperature Setpoint Node Name
  0.1;              !- Parasitic Electric Load W";
            return string.Format(
                template,
                name,
                reheatCoil["Availability Schedule Name"].Value,
                reheatCoil["Air Inlet Node Name"].Value,
                reheatCoil["Air Outlet Node Name"].Value,
                coolingCoilType,
                coolingCoilName,
                reheatCoil["Temperature Setpoint Node Name"].Value);
        }

        private void AddDesuperheaterToUnitary(IdfObject unitary)
        {
            string unitaryName = unitary["Name"].Value;

            string reheatCoilType = unitary["Supplemental Heating Coil Object Type"].Value;
            string reheatCoilName = unitary["Supplemental Heating Coil Name"].Value;
            string coolingCoilType = unitary["Cooling Coil Object Type"].Value;
            string coolingCoilName = unitary["Cooling Coil Name"].Value;

            if (string.IsNullOrEmpty(reheatCoilType))
            {
                MessageBox.Show("Skipping unit: " + unitaryName + ", reheat coil is not included." +
                                "\nMake sure that the unit uses CoolReheat humidity control.");
                return;
            }

            IdfObject reheatCoil = FindObject(reheatCoilType, reheatCoilName);

            string desuperheaterName = unitaryName + " Desuperheater Coil";
            string desuperheater = GetDesuperheaterCoil(desuperheaterName, coolingCoilType, coolingCoilName, reheatCoil);

            unitary["Supplemental Heating Coil Object Type"].Value = "Coil:Heating:Desuperheater";
            unitary["Supplemental Heating Coil Name"].Value = desuperheaterName;

            Reader.Load(desuperheater);
            Reader.Remove(reheatCoil);
        }

        public override void BeforeEnergySimulation()
        {
            Reader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            foreach (IdfObject unitary in Reader["AirLoopHVAC:UnitarySystem"])
            {
                string name = unitary[0].Value;
                if (name.ToLower().Contains("desuperheater"))
                {
                    AddDesuperheaterToUnitary(unitary);
                }
            }
            Reader.Save();
        }
    }
}
