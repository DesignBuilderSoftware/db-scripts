/*
Replace loop supply side "Scheduled" setpoint manager with the SetpointManager:ReturnTemperature:* equivalent.

SetpointManager:Scheduled objects with:
- "CHW Return" key will be replaced by "SetpointManager:ReturnTemperature:ChilledWater"
- "HW Return" key will be replaced by "SetpointManager:ReturnTemperature:HotWater"

The return setpoint manager follows the original suply side schedule.
*/

using System;
using System.Runtime;
using System.Linq;
using System.Windows.Forms;

using EpNet;
using DB.Extensibility.Contracts;


namespace DB.Extensibility.Scripts
{
    public class ApplyLoopReturnSpm : ScriptBase, IScript
    {
        private string hwReturnSpmTemplate = @"SetpointManager:ReturnTemperature:HotWater,
  {0},                       !- Name
  {1},                       !- Plant Loop Supply Outlet Node
  {2},                       !- Plant Loop Supply Inlet Node
  57.0,                      !- Minimum Supply Temperature Setpoint
  60.0,                      !- Maximum Supply Temperature Setpoint
  ReturnTemperatureSetpoint, !- Return Temperature Setpoint Input Type
  ,                          !- Return Temperature Setpoint Constant Value
  ;                          !- Return Temperature Setpoint Schedule Name";

        private string chwReturnSpmTemplate = @"SetpointManager:ReturnTemperature:ChilledWater,
  {0},                       !- Name
  {1},                       !- Plant Loop Supply Outlet Node
  {2},                       !- Plant Loop Supply Inlet Node
  7.0,                       !- Minimum Supply Temperature Setpoint
  10.0,                      !- Maximum Supply Temperature Setpoint
  ReturnTemperatureSetpoint, !- Return Temperature Setpoint Input Type
  ,                          !- Return Temperature Setpoint Constant Value
  ;                          !- Return Temperature Setpoint Schedule Name";


        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // Define look-up keys, check is case agnostic
            const string chwKey = "chw return";
            const string hwKey = "hw return";

            ApplyReturnSpms(idfReader, chwKey, hwKey);

            idfReader.Save();
        }

        private IdfObject FindObject(IdfReader reader, string objectType, string objectName)
        {
            try
            {
                return reader[objectType].First(c => c[0] == objectName);
            }
            catch (Exception e)
            {
                throw new MissingFieldException(String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        private void ApplyReturnSpms(IdfReader reader, string chwKey, string hwKey)
        {
            var spms = reader["SetpointManager:Scheduled"];
            foreach (var spm in spms)
            {
                string name = spm[0].Value;
                string template = "";
                if (name.IndexOf(chwKey, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    template = chwReturnSpmTemplate;
                } else if (name.IndexOf(hwKey, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    template = hwReturnSpmTemplate;
                }
                else
                {
                    continue;
                }
                string returnSpm = GetReturnSpm(reader, spm, template);
                reader.Load(returnSpm);
            }
        }

        private string GetReturnSpm(IdfReader reader, IdfObject spm, string template)
        {   
            string nodeListName = spm[3].Value;
            IdfObject nodeList = FindObject(reader, "NodeList", nodeListName);
            string supplySideOutletNodeName = nodeList[1].Value;

            string supplySideInletNodeName = supplySideOutletNodeName.Replace("Outlet", "Inlet");
            nodeList[1].Value = supplySideInletNodeName;
            return string.Format(template, "Main " + spm[0].Value, supplySideOutletNodeName, supplySideInletNodeName);
        }
    }
}