/*
Apply Desuperheater Reheat Coil to AirLoopHVAC:UnitaryHeatCool Systems

Purpose:
This DesignBuilder C# script modifies the EnergyPlus IDF to replace the existing reheat coil on selected AirLoopHVAC:UnitaryHeatCool units
with a Coil:Heating:Desuperheater that reclaims heat from the unit’s cooling coil.

Main Steps:
1) Scan all AirLoopHVAC:UnitaryHeatCool objects in the IDF
2) Select target units by name (must contain a configurable substring, by default set as "desuperheater")
3) For each target unit:
  - Read the existing reheat coil fields (used as a placeholder for nodes/schedule/setpoint node)
  - Create a new Coil:Heating:Desuperheater IDF object using the placeholder reheat coil data
  - Update the unitary “Reheat Coil Object Type/Name” fields to point to the new desuperheater coil
4) Save the updated IDF

How to Use:

Configuration
- Name your target unitary systems so their Name contains the selection substring (default: "desuperheater").
- Define desuperheater coil efficiency (see desuperheaterCoilTemplate).
- Availability schedule, temperature setpoint, and node connections are taken from the existing reheat coil (placeholder).

Prerequisites / Placeholders
- Eligible AirLoopHVAC:UnitaryHeatCool objects must have the trigger substring ("desuperheater") in the name
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
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        // Selection rule for target units (case-insensitive substring match on AirLoopHVAC:UnitaryHeatCool Name).
        private const string SelectionSubstring = "desuperheater";

        private IdfReader idf;

        private IdfObject FindObject(string objectType, string objectName)
        {
            try
            {
                return idf[objectType].First(c => c[0] == objectName);
            }
            catch (Exception)
            {
                throw new MissingFieldException(
                    string.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        // Generates the IDF text for a Coil:Heating:Desuperheater object.
        // Availability schedule, temperature setpoint, and node connections are taken from the existing reheat coil (placeholder).
        private string GetDesuperheaterCoilIdfText(
            string desuperheaterCoilName,
            string coolingCoilType,
            string coolingCoilName,
            IdfObject placeholderReheatCoil)
        {
            // USER CONFIGURATION: Define the efficiency of the Desuperheater (default is set to 0.3)
            string desuperheaterCoilTemplate = @"
Coil:Heating:Desuperheater,
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
                desuperheaterCoilTemplate,
                desuperheaterCoilName,
                placeholderReheatCoil["Availability Schedule Name"].Value,
                placeholderReheatCoil["Air Inlet Node Name"].Value,
                placeholderReheatCoil["Air Outlet Node Name"].Value,
                coolingCoilType,
                coolingCoilName,
                placeholderReheatCoil["Temperature Setpoint Node Name"].Value);
        }

        // Replaces the unitary reheat coil reference with a new Coil:Heating:Desuperheater and loads the new object into the IDF.
        private void AddDesuperheaterToUnitaryHeatCool(IdfObject unitaryHeatCool)
        {
            string unitaryName = unitaryHeatCool["Name"].Value;

            string reheatCoilType = unitaryHeatCool["Reheat Coil Object Type"].Value;
            string reheatCoilName = unitaryHeatCool["Reheat Coil Name"].Value;
            string coolingCoilType = unitaryHeatCool["Cooling Coil Object Type"].Value;
            string coolingCoilName = unitaryHeatCool["Cooling Coil Name"].Value;

            if (string.IsNullOrEmpty(reheatCoilType))
            {
                MessageBox.Show(
                    "Skipping unit: " + unitaryName + ", reheat coil is not included." +
                    "\nMake sure that the unit uses CoolReheat humidity control.");
                return;
            }

            // Placeholder reheat coil provides nodes/schedule/setpoint node for the new desuperheater coil.
            IdfObject placeholderReheatCoil = FindObject(reheatCoilType, reheatCoilName);

            string desuperheaterCoilName = unitaryName + " Desuperheater Coil";
            string desuperheaterCoilIdfText = GetDesuperheaterCoilIdfText(
                desuperheaterCoilName,
                coolingCoilType,
                coolingCoilName,
                placeholderReheatCoil);

            // Update the unitary to point to the new desuperheater coil object.
            unitaryHeatCool["Reheat Coil Object Type"].Value = "Coil:Heating:Desuperheater";
            unitaryHeatCool["Reheat Coil Name"].Value = desuperheaterCoilName;

            idf.Load(desuperheaterCoilIdfText);
        }

        public override void BeforeEnergySimulation()
        {
            idf = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            foreach (IdfObject unitaryHeatCool in idf["AirLoopHVAC:UnitaryHeatCool"])
            {
                string unitaryName = unitaryHeatCool[0].Value;

                // Target units are identified by substring match in the unit name (case-insensitive).
                if (unitaryName.IndexOf(SelectionSubstring, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    AddDesuperheaterToUnitaryHeatCool(unitaryHeatCool);
                }
            }

            idf.Save();
        }
    }
}