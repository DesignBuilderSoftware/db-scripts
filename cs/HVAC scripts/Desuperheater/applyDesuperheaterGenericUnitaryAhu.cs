/*
Apply Desuperheater Coil to Generic Unitary System Objects

Purpose:
This DesignBuilder C# script runs before the EnergyPlus simulation and modifies the IDF to replace a unitary
system’s supplemental (reheat) heating coil with a Coil:Heating:Desuperheater for eligible systems.

Main Steps:
1) Find all AirLoopHVAC:UnitarySystem objects whose Name contains a configurable substring (default: "desuperheater").
2) For each match:
   - Read the unitary system’s Supplemental Heating Coil and Cooling Coil references
   - Build a new Coil:Heating:Desuperheater object using the reheat coil nodes/schedule and the cooling coil as heat source
   - Update the unitary system to reference the new desuperheater coil
   - Insert the new desuperheater coil object into the IDF
   - Remove the original reheat coil object from the IDF

How to Use:

Configuration
- Name your target AirLoopHVAC:UnitarySystem objects so their Name contains the trigger substring.
  Default trigger substring is: "desuperheater" (case-insensitive).
- Define desuperheater coil efficiency (see desuperheaterCoilIdfTemplate).
- Availability schedule, temperature setpoint, and node connections are taken from the existing reheat coil (placeholder).

Prerequisites / Placeholders
- Eligible AirLoopHVAC:UnitarySystem objects must have the trigger substring ("desuperheater") in the name
- The unit must include CoolReheat humidity control so that a supplemental (reheat) coil is present and referenced.
    This setting can be controled in 'Unitary System settings > Control > Dehumidification control type' in DB model.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. 
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Linq;
using System.Windows.Forms;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class ApplyDesuperheaterToUnitarySystems : ScriptBase, IScript
    {
        private IdfReader idfReader;

        private IdfObject FindIdfObject(string objectType, string objectName)
        {
            try
            {
                return idfReader[objectType].First(o => o[0] == objectName);
            }
            catch
            {
                throw new MissingFieldException(
                    string.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        // Generates the IDF text for a Coil:Heating:Desuperheater object.
        // Availability schedule, temperature setpoint, and node connections are taken from the existing reheat coil (placeholder).
        private string BuildDesuperheaterCoilIdf(
            string desuperheaterCoilName,
            string coolingCoilObjectType,
            string coolingCoilObjectName,
            IdfObject existingReheatCoil)
        {
            // USER CONFIGURATION: Define the efficiency of the Desuperheater (default is set to 0.3)
            string desuperheaterCoilIdfTemplate = @"
Coil:Heating:Desuperheater,
  {0},              !- Coil Name
  {1},              !- Availability Schedule
  {2},              !- Coil Air Inlet Node Name
  {3},              !- Coil Air Outlet Node Name
  {4},              !- Heating Source Type
  {5},              !- Heating Source Name
  {6},              !- Coil Temperature Setpoint Node Name
  0.1;              !- Parasitic Electric Load W";

            return string.Format(
                desuperheaterCoilIdfTemplate,
                desuperheaterCoilName,
                existingReheatCoil["Availability Schedule Name"].Value,
                existingReheatCoil["Air Inlet Node Name"].Value,
                existingReheatCoil["Air Outlet Node Name"].Value,
                coolingCoilObjectType,
                coolingCoilObjectName,
                existingReheatCoil["Temperature Setpoint Node Name"].Value);
        }

        // Replaces the unitary system's "Supplemental Heating Coil" reference with a new desuperheater coil
        private void ApplyDesuperheaterToUnitarySystem(IdfObject unitarySystem)
        {
            string unitarySystemName = unitarySystem["Name"].Value;

            string reheatCoilObjectType = unitarySystem["Supplemental Heating Coil Object Type"].Value;
            string reheatCoilObjectName = unitarySystem["Supplemental Heating Coil Name"].Value;

            string coolingCoilObjectType = unitarySystem["Cooling Coil Object Type"].Value;
            string coolingCoilObjectName = unitarySystem["Cooling Coil Name"].Value;

            if (string.IsNullOrEmpty(reheatCoilObjectType))
            {
                MessageBox.Show(
                    "Skipping unit: " + unitarySystemName +
                    "\nSupplemental (reheat) coil is not included." +
                    "\nMake sure that the unit uses CoolReheat humidity control.");
                return;
            }

            // Placeholder reheat coil object is used to copy schedule + node connections
            IdfObject existingReheatCoil = FindIdfObject(reheatCoilObjectType, reheatCoilObjectName);

            string desuperheaterCoilName = unitarySystemName + " Desuperheater Coil";
            string desuperheaterCoilIdf = BuildDesuperheaterCoilIdf(
                desuperheaterCoilName,
                coolingCoilObjectType,
                coolingCoilObjectName,
                existingReheatCoil);

            // Update unitary system to reference the new desuperheater coil as the supplemental heating coil
            unitarySystem["Supplemental Heating Coil Object Type"].Value = "Coil:Heating:Desuperheater";
            unitarySystem["Supplemental Heating Coil Name"].Value = desuperheaterCoilName;

            idfReader.Load(desuperheaterCoilIdf);
            idfReader.Remove(existingReheatCoil);
        }

        public override void BeforeEnergySimulation()
        {
            idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // Configuration: select unitary systems by substring match in Name (case-insensitive).
            const string nameFilterSubstring = "desuperheater";

            foreach (IdfObject unitarySystem in idfReader["AirLoopHVAC:UnitarySystem"])
            {
                string unitarySystemName = unitarySystem[0].Value;

                // Only apply to systems whose name contains the configured substring.
                if (unitarySystemName.IndexOf(nameFilterSubstring, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    ApplyDesuperheaterToUnitarySystem(unitarySystem);
                }
            }

            idfReader.Save();
        }
    }
}