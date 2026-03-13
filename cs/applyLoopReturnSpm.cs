/*
Apply Loop Return Setpoint Manager Script (Plant Loops)

This DesignBuilder C# script replaces loop supply-side SetpointManager:Scheduled objects (with specific name keys) 
by creating the equivalent SetpointManager:ReturnTemperature:* object for plant loops.

Purpose
1) Identify SetpointManager:Scheduled objects whose name contains:
  - "CHW Return"  -> create SetpointManager:ReturnTemperature:ChilledWater
  - "HW Return"   -> create SetpointManager:ReturnTemperature:HotWater
2) For each matched scheduled SPM:
  - Read its NodeList reference
  - Extract the supply outlet node name from the NodeList
  - Infer supply inlet node name by replacing "Outlet" with "Inlet"
  - Create and load the appropriate ReturnTemperature setpoint manager object into the IDF
  - (Side-effect) Update NodeList to point to the inferred inlet node (see notes in code)

How to Use

Configuration
- Matching keys are controlled by the constants:
  chilledWaterKey = "CHW Return"
  hotWaterKey     = "HW Return"
  (matching is case-insensitive)
- Default minimum/maximum supply temperature limits are defined in the templates:
  - CHW: 7°C min, 10°C max
  - HW : 57°C min, 60°C max
  Adjust these values in the templates if needed.

Prerequisites (required placeholders)

- The base model must include SetpointManager:Scheduled objects used as placeholders,
  with Names containing either "CHW Return" or "HW Return".

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Linq;

using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class ApplyLoopReturnSpm : ScriptBase, IScript
    {
        // ----------------------------
        // USER CONFIGURATION SECTION
        // ----------------------------
        // IDF templates for the two possible return-temperature setpoint managers.
        // Adjust minimum/maximum supply temperature limits if needed.
        private readonly string hotWaterReturnSpmTemplate = @"
SetpointManager:ReturnTemperature:HotWater,
  {0},                       !- Name
  {1},                       !- Plant Loop Supply Outlet Node
  {2},                       !- Plant Loop Supply Inlet Node
  57.0,                      !- Minimum Supply Temperature Setpoint
  60.0,                      !- Maximum Supply Temperature Setpoint
  ReturnTemperatureSetpoint, !- Return Temperature Setpoint Input Type
  ,                          !- Return Temperature Setpoint Constant Value
  ;                          !- Return Temperature Setpoint Schedule Name";

        private readonly string chilledWaterReturnSpmTemplate = @"
SetpointManager:ReturnTemperature:ChilledWater,
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
            IdfReader idf = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // Lookup keys used to decide whether a SetpointManager:Scheduled is a placeholder (case-insensitive).
            const string chilledWaterKey = "CHW Return";
            const string hotWaterKey = "HW Return";

            ApplyReturnSetpointManagers(idf, chilledWaterKey, hotWaterKey);

            idf.Save();
        }

        private static IdfObject FindObject(IdfReader reader, string objectType, string objectName)
        {
            try
            {
                return reader[objectType].First(o => o[0] == objectName);
            }
            catch
            {
                throw new MissingFieldException(
                    string.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        private void ApplyReturnSetpointManagers(IdfReader reader, string chilledWaterKey, string hotWaterKey)
        {
            // Iterate through all scheduled setpoint managers and select those with matching keys.
            var scheduledSpms = reader["SetpointManager:Scheduled"];

            foreach (var scheduledSpm in scheduledSpms)
            {
                string scheduledSpmName = scheduledSpm[0].Value;
                string returnSpmTemplate = null;

                if (scheduledSpmName.IndexOf(chilledWaterKey, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    returnSpmTemplate = chilledWaterReturnSpmTemplate;
                }
                else if (scheduledSpmName.IndexOf(hotWaterKey, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    returnSpmTemplate = hotWaterReturnSpmTemplate;
                }
                else
                {
                    continue;
                }

                // Build the ReturnTemperature setpoint manager IDF snippet and load it into the model.
                string returnSpmIdfText = BuildReturnSpm(reader, scheduledSpm, returnSpmTemplate);
                reader.Load(returnSpmIdfText);
            }
        }

        private static string BuildReturnSpm(IdfReader reader, IdfObject scheduledSpm, string returnSpmTemplate)
        {
            string nodeListName = scheduledSpm[3].Value;
            IdfObject nodeList = FindObject(reader, "NodeList", nodeListName);
            string supplyOutletNodeName = nodeList[1].Value;

            string supplyInletNodeName = supplyOutletNodeName.Replace("Outlet", "Inlet");
            nodeList[1].Value = supplyInletNodeName;
            string returnSpmName = "Main " + scheduledSpm[0].Value;

            return string.Format(returnSpmTemplate, returnSpmName, supplyOutletNodeName, supplyInletNodeName);
        }
    }
}