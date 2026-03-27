/*
Targeted Node Output Variable Request Script

Purpose:
This DesignBuilder C# script adds Output:Variable objects for a limited set of node names to help avoid extremely 
large output files that can occur when reporting variables for all nodes in large models.

Main Steps:
1) Collect outlet node names from AirLoopHVAC, PlantLoop and CondenserLoop objects.
   - If a node name ends with "List", it is assumed to reference a NodeList object, and the script uses the first node entry in that NodeList.
2) For each requested variable name, add Output:Variable objects for each collected node.
3) Save the modified IDF before the EnergyPlus simulation runs.

How to Use:

Configuration
- requestedVariables:
  Add/remove entries using the exact EnergyPlus variable names you want to request. For example:
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
- reportingFrequency:
  Set to a valid Output:Variable reporting frequency string.

Prerequisites / Placeholders
Base model must contain the required objects/fields:
- AirLoopHVAC objects (Supply Side Outlet Node Names)
- PlantLoop objects (Plant Side Outlet Node Name)
- CondenserLoop objects (Condenser Side Outlet Node Name)
If any of these fields reference a NodeList, the NodeList object must exist in the IDF.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. 
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System.Collections.Generic;
using System.Linq;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    // Adds Output:Variable requests for specific node variables and specific outlet nodes only.
    public class AddTargetedNodeOutputVariables : ScriptBase, IScript
    {
        // Collect node names from a given IDF object type and field name.
        // If a value ends with "List", it is treated as a NodeList reference and the first node entry is used.
        private List<string> CollectOutletNodeNames(IdfReader idfReader, string idfObjectType, string nodeFieldName)
        {
            var nodeNames = new List<string>();
            IEnumerable<IdfObject> objectsOfType = idfReader[idfObjectType];

            foreach (IdfObject obj in objectsOfType)
            {
                string nodeOrListName = obj[nodeFieldName];

                if (nodeOrListName.EndsWith("List"))
                {
                    IdfObject nodeList = idfReader["NodeList"].First(item => item[0] == nodeOrListName);
                    nodeOrListName = nodeList[1];
                }

                nodeNames.Add(nodeOrListName);
            }

            return nodeNames;
        }

        private void AddOutputVariable(IdfReader idfReader, string nodeName, string variableName, string reportingFrequency)
        {
            string outputVariableObject = string.Format(
                "Output:Variable, {0}, {1}, {2};",
                nodeName,
                variableName,
                reportingFrequency);

            idfReader.Load(outputVariableObject);
        }

        private void AddVariablesForNodes(IdfReader idfReader, List<string> nodeNames, string variableName, string reportingFrequency)
        {
            foreach (string nodeName in nodeNames)
            {
                AddOutputVariable(idfReader, nodeName, variableName, reportingFrequency);
            }
        }

        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // Collect outlet nodes from the three loop types (Air, Plant, Condenser)
            List<string> airLoopOutletNodes = CollectOutletNodeNames(idfReader, "AirLoopHVAC", "Supply Side Outlet Node Names");
            List<string> plantLoopOutletNodes = CollectOutletNodeNames(idfReader, "PlantLoop", "Plant Side Outlet Node Name");
            List<string> condenserLoopOutletNodes = CollectOutletNodeNames(idfReader, "CondenserLoop", "Condenser Side Outlet Node Name");

            // ----------------------------
            // USER CONFIGURATION SECTION
            // ----------------------------
            // Add additional EnergyPlus Output:Variable names to the list below as required.
            List<string> requestedVariables = new List<string>
            {
                "System Node Temperature",
                "System Node Mass Flow Rate"
            };

            // Add desired reporting frequency 
            const string reportingFrequency = "hourly";

            foreach (string variableName in requestedVariables)
            {
                AddVariablesForNodes(idfReader, plantLoopOutletNodes, variableName, reportingFrequency);
                AddVariablesForNodes(idfReader, airLoopOutletNodes, variableName, reportingFrequency);
                AddVariablesForNodes(idfReader, condenserLoopOutletNodes, variableName, reportingFrequency);
            }

            idfReader.Save();
        }
    }
}