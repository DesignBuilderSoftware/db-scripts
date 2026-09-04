/*
Sets the heat recovery leaving water temperature setpoint node for an electric chiller.

Purpose:
Electric chillers with heat recovery (Chiller:Electric:EIR and Chiller:Electric:ReformulatedEIR) can control
the temperature of the water leaving the heat recovery bundle. To do so, EnergyPlus needs the field
"Heat Recovery Leaving Temperature Setpoint Node Name" to point at a node (or NodeList) that carries a
setpoint placed there by a SetpointManager. DesignBuilder does not expose this field in the Detailed HVAC
interface, so it is written blank into the IDF and the heat recovery leaving temperature is left uncontrolled.
This script writes the configured node name into that field before the simulation starts.

Main Steps:
1) Locate the chiller object by type and name.
2) Writes the configured setpoint node (or NodeList) name into the chiller's
   "Heat Recovery Leaving Temperature Setpoint Node Name" field.
3) Saves the modified IDF so the simulation runs with the updated chiller object.

How to Use:

Configuration
- chillerName: the chiller name exactly as it appears in the DesignBuilder interface (e.g., "Chiller").
- heatRecoverySetpointNodeName: the Node or NodeList holding the heat recovery leaving water setpoint
  (default: "HR Loop Setpoint Manager Node List").
- Chiller object type: call SetHeatRecoverySetpointNodeEirChiller for Chiller:Electric:EIR, or
  SetHeatRecoverySetpointNodeReformulatedEirChiller for Chiller:Electric:ReformulatedEIR. Only one of the
  two calls should be active; the other is left commented out in the USER CONFIGURATION SECTION.

Prerequisites
- The model must contain an electric chiller with heat recovery enabled, so that the IDF object already has
  its "Heat Recovery Inlet Node Name" and "Heat Recovery Outlet Node Name" fields populated.
- A SetpointManager must already place a temperature setpoint on the node named in heatRecoverySetpointNodeName. 
  This script does not create the setpoint manager or the NodeList.

Notes:
- Only one chiller is handled per call. For several chillers, repeat the call once per chiller name, or wrap
  it in a loop over an array of names in the USER CONFIGURATION SECTION.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. 
Users are responsible for validating all outputs and ensuring the script meets their specific modelling requirements.
*/

using System.Runtime;
using System;
using System.Linq;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        // Hook Point: Executes before the EnergyPlus simulation starts
        public override void BeforeEnergySimulation()
        {
            // 1. Initialize the IdfReader to modify the simulation input
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // ---------------------------
            // USER CONFIGURATION SECTION
            // ---------------------------

            // Chiller name exactly as written in the IDF (case-sensitive)
            string chillerName = "Chiller";

            // Node or NodeList carrying the heat recovery leaving water setpoint
            string heatRecoverySetpointNodeName = "HR Loop Setpoint Manager Node List";

            // 2. Apply the setpoint node to the chiller.
            //    Keep active only the line matching the chiller object type used in the model.
            SetHeatRecoverySetpointNodeEirChiller(idfReader, chillerName, heatRecoverySetpointNodeName);
            // SetHeatRecoverySetpointNodeReformulatedEirChiller(idfReader, chillerName, heatRecoverySetpointNodeName);

            idfReader.Save();
        }

        // Returns the first object of the given type whose Name field (index 0) matches objectName.
        private IdfObject FindObject(IdfReader idfReader, string objectType, string objectName)
        {
            try
            {
                return idfReader[objectType].First(idfObject => idfObject[0] == objectName);
            }
            catch (Exception exception)
            {
                throw new Exception(
                    String.Format("Cannot find object: {0}, type: {1}", objectName, objectType),
                    exception);
            }
        }

        // Chiller:Electric:ReformulatedEIR variant.
        private void SetHeatRecoverySetpointNodeReformulatedEirChiller(IdfReader idfReader, string chillerName, string setpointNodeName)
        {
            IdfObject chiller = FindObject(idfReader, "Chiller:Electric:ReformulatedEIR", chillerName);
            chiller["Heat Recovery Leaving Temperature Setpoint Node Name"].Value = setpointNodeName;
        }

        // Chiller:Electric:EIR varian
        private void SetHeatRecoverySetpointNodeEirChiller(IdfReader idfReader, string chillerName, string setpointNodeName)
        {
            IdfObject chiller = FindObject(idfReader, "Chiller:Electric:EIR", chillerName);
            chiller["Heat Recovery Leaving Temperature Setpoint Node Name"].Value = setpointNodeName;
        }
    }
}