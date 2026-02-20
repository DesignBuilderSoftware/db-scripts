/*
Reporting all node outputs can be difficult as output files tend to be enormous due to
huge amount of data (especially for bigger models).

This script adds requested output ariables only for defined nodes:
    "Supply Side Outlet Node Names"
    "Plant Side Outlet Node Name"
    "Condenser Side Outlet Node Name"

Default variables are:
    System Node Temperature
    System Node Mass Flow Rate

Additional outputs (below) can be added to 'variables' list:
    System Node Temperature
    System Node Mass Flow Rate
    System Node Humidity Ratio
    System Node Setpoint Temperature
    System Node Setpoint High Temperature
    System Node Setpoint Low Temperature
    System Node Setpoint Humidity Ratio
    System Node Setpoint Minimum Humidity Ratio
    System Node Setpoint Maximum Humidity Ratio
    System Node Relative Humidity
    System Node Pressure
    System Node Standard Density Volume Flow Rate
    System Node Current Density Volume Flow Rate
    System Node Current Density
    System Node Specific Heat
    System Node Enthalpy
    System Node Minimum Temperature
    System Node Maximum Temperature
    System Node Minimum Limit Mass Flow Rate
    System Node Maximum Limit Mass Flow Rate
    System Node Minimum Available Mass Flow Rate
    System Node Maximum Available Mass Flow Rate
    System Node Setpoint Mass Flow Rate
    System Node Requested Mass Flow Rate
*/

using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        private List<string> FindNodes(IdfReader idfReader, string objectName, string fieldName)
        {
            List<string> nodes = new List<string>();
            IEnumerable<IdfObject> idfObjects = idfReader[objectName];

            foreach (IdfObject idfObject in idfObjects)
            {
                string nodeName = idfObject[fieldName];

                if (nodeName.EndsWith("List"))
                {
                    IdfObject nodeList = idfReader["NodeList"].First(item => item[0] == nodeName);
                    nodeName = nodeList[1];
                }
                nodes.Add(nodeName);
            }
            return nodes;
        }

        private void AddOuputVariable(IdfReader idfReader, string key, string name, string frequency)
        {
            string outputVariable = string.Format("Output:Variable, {0}, {1}, {2};", key, name, frequency);
            idfReader.Load(outputVariable);
        }

        private void AddNodeVariables(IdfReader idfReader, List<string> nodes, string name, string frequency)
        {
            foreach (string node in nodes)
            {
                AddOuputVariable(idfReader, node, name, frequency);
            }
        }

        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            List<string> airLoopNodes = FindNodes(idfReader, "AirLoopHVAC", "Supply Side Outlet Node Names");
            List<string> plantLoopNodes = FindNodes(idfReader, "PlantLoop", "Plant Side Outlet Node Name");
            List<string> condenserLoopNodes = FindNodes(idfReader, "CondenserLoop", "Condenser Side Outlet Node Name");

            // Request node variables (listed in description above) by adding them into list below
            List<string> variables = new List<string> { "System Node Temperature", "System Node Mass Flow Rate" };
            const string frequency = "hourly";

            foreach (string variable in variables)
            {
                AddNodeVariables(idfReader, plantLoopNodes, variable, frequency);
                AddNodeVariables(idfReader, airLoopNodes, variable, frequency);
                AddNodeVariables(idfReader, condenserLoopNodes, variable, frequency);
            }

            idfReader.Save();
        }
    }
}