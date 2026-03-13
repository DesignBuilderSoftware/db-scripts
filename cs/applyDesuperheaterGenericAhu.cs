/*
Replace Electric Heating Coils with Desuperheater Coils (for Generic AHUs)

This DesignBuilder C# script scans EnergyPlus Branch objects and replaces
Coil:Heating:Electric entries with Coil:Heating:Desuperheater for selected AHUs.

Purpose
1) Identify target Branch objects by name filters:
   - Branch Name contains "desuperheater" and "AHU Main Branch"
2) For each target Branch:
   - Confirm the Branch includes a CoilSystem:Cooling:DX component (DX coil system).
   - Resolve the actual DX cooling coil referenced by the DX coil system.
   - Replace the first Coil:Heating:Electric found in the Branch with Coil:Heating:Desuperheater.
   - Create a new Coil:Heating:Desuperheater object using the replaced electric coil’s nodes/schedule,
     and referencing the resolved DX cooling coil.
   - Remove the original Coil:Heating:Electric object from the IDF.
3) Save the modified IDF before the EnergyPlus simulation runs.

How to Use

Configuration
- Target selection is controlled by Branch Name text matching:
  - Must contain: "desuperheater" and "AHU Main Branch"
- Define desuperheater coil efficiency (see AddDesuperheaterCoil()).

Prerequisites (required placeholders)

Target Branches must include:
- A "CoilSystem:Cooling:DX" component (the script uses this to find the referenced DX cooling coil).
- A "Coil:Heating:Electric" component to be replaced by the Desuperheater. It also provides:
    - Availability schedule (as defined in the model)
    - Temperature setpoint node name (as defined in the model)
    - Inlet / outlet node names (automatically obtained)

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class ReplaceElectricHeatingWithDesuperheaterCoils : ScriptBase, IScript
    {
        // Returns the first object of the given type with Name == objectName.
        private IdfObject GetObjectByName(IdfReader reader, string objectType, string objectName)
        {
            return reader[objectType].First(c => c[0] == objectName);
        }

        // Entry point for IDF edits: find all candidate Branch objects and update them.
        private void ReplaceCoilsInTargetBranches(IdfReader idfReader)
        {
            IEnumerable<IdfObject> branches = idfReader["Branch"];

            foreach (IdfObject branch in branches)
            {
                // Target selection based on Branch Name text matching.
                string branchName = branch[0].Value;

                if (branchName.Contains("desuperheater") && branchName.Contains("AHU Main Branch"))
                {
                    ReplaceCoilInBranchIfEligible(idfReader, branch);
                }
            }
        }

    
        // Replace Coil:Heating:Electric with Coil:Heating:Desuperheater (if CoilSystem:Cooling:DX exists).
        private void ReplaceCoilInBranchIfEligible(IdfReader reader, IdfObject branch)
        {
            const string DesuperheaterObjectType = "Coil:Heating:Desuperheater";
            const string ElectricHeatingCoilObjectType = "Coil:Heating:Electric";
            const string DxCoilSystemObjectType = "CoilSystem:Cooling:DX";

            IdfObject dxCoolingCoil = new IdfObject();
            bool hasDxCoilSystem = false;

            // Branch objects typically list equipment in pairs
            foreach (int i in Enumerable.Range(1, (branch.Count - 2)))
            {
                Field objectTypeField = branch.Fields[i];
                Field objectNameField = branch.Fields[i + 1];

                // Confirm if the branch contains a DX cooling coil system and resolve its cooling coil.
                if (objectTypeField.Value.Contains(DxCoilSystemObjectType))
                {
                    MessageBox.Show("DX Coil System included " + branch[0]);

                    hasDxCoilSystem = true;

                    IdfObject dxCoilSystem = GetObjectByName(reader, DxCoilSystemObjectType, objectNameField.Value);

                    dxCoolingCoil = GetObjectByName(
                        reader,
                        dxCoilSystem["Cooling Coil Object Type"],
                        dxCoilSystem["Cooling Coil Name"]);
                }

                // Once DX is confirmed, replace the electric heating coil object type in the branch list.
                if (hasDxCoilSystem
                    && objectTypeField.Comment.ToLower().Contains("object type")
                    && objectTypeField.Value.ToLower() == ElectricHeatingCoilObjectType.ToLower())
                {
                    IdfObject electricHeatingCoil = GetObjectByName(reader, ElectricHeatingCoilObjectType, objectNameField.Value);

                    // Replace the object type in the Branch equipment list (object name stays the same).
                    objectTypeField.Value = DesuperheaterObjectType;

                    AddDesuperheaterCoil(reader, electricHeatingCoil, dxCoolingCoil);

                    MessageBox.Show(
                        "Replacing electric heating coil: " + electricHeatingCoil[0] + " with desuperheater coil.");

                    reader.Remove(electricHeatingCoil);

                    break;
                }
            }
        }

        private void AddDesuperheaterCoil(IdfReader reader, IdfObject electricCoil, IdfObject dxCoil)
        {
            // Placeholder electric coil fields are used to populate the new desuperheater coil.
            string name = electricCoil[0].Value;
            string availabilityScheduleName = electricCoil[1].Value;
            string airInletNodeName = electricCoil[4].Value;
            string airOutletNodeName = electricCoil[5].Value;
            string temperatureSetpointNodeName = electricCoil[6].Value;

            // The desuperheater coil references the DX cooling coil (object type/class + name).
            string dxCoolingCoilObjectType = dxCoil.IdfClass;
            string dxCoolingCoilName = dxCoil[0].Value;

            double desuperheaterEfficiency = 0.3; // USER CONFIGURATION: Define the efficiency of the Desuperheater (default is set to 0.3)

            string desuperheaterIdf = String.Format(
                "Coil:Heating:Desuperheater,{0},{1},{2},{3},{4},{5},{6},{7},0;",
                name,
                availabilityScheduleName,
                desuperheaterEfficiency,
                airInletNodeName,
                airOutletNodeName,
                dxCoolingCoilObjectType,
                dxCoolingCoilName,
                temperatureSetpointNodeName);

            reader.Load(desuperheaterIdf);
        }

        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            ReplaceCoilsInTargetBranches(idfReader);

            idfReader.Save();
        }
    }
}