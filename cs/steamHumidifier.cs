/*
Replace Electric Steam Humidifiers with Gas Steam Humidifiers

Purpose
This DesignBuilder C# script modifies the EnergyPlus IDF by replacing Humidifier:Steam:Electric 
objects with Humidifier:Steam:Gas objects on selected air loops.

Main Steps
1) Find Branch objects that match the naming convention:
   - Branch name contains both "steamgas" and "AHU Main Branch"
2) Within matching Branch objects, locate equipment entries where:
   - The field comment contains an object type "Humidifier:Steam:Electric"
3) For each matched humidifier:
   - Update the Branch reference from Humidifier:Steam:Electric to Humidifier:Steam:Gas
   - Create and load a new Humidifier:Steam:Gas object using key fields from the electric humidifier
   - Remove the original Humidifier:Steam:Electric object from the IDF
4) Add Output:Variable requests for humidifier natural gas rate.
5) Save the updated IDF.

How to Use

Configuration
- The script only updates Branch objects where the Branch name contains both "steamgas" and "AHU Main Branch"
  To change the selection rule, edit the string checks in UpdateBranches().
- Gas humidifier assumptions:
  - Thermal efficiency is set by default to 0.9 in GetHumidifierIdfText().
  - Other key fields are copied from the original electric humidifier object by index.
  If you need different default gas humidifier inputs, edit GetHumidifierIdfText().
- Output "Humidifier NaturalGas Rate" is added by default for hourly and Runperiod timestamps (edit for other periods).

Prerequisites / Placeholders
- The model must already include one or more Humidifier:Steam:Electric objects.
- Branch naming must include the required keywords ("steamgas" and "AHU Main Branch") to be modified.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script.
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        private IdfObject FindObject(IdfReader reader, string objectType, string objectName)
        {
            return reader[objectType].First(c => c[0] == objectName);
        }

        private void UpdateBranches(IdfReader idfReader)
        {
            IEnumerable<IdfObject> branches = idfReader["Branch"];
            foreach (IdfObject branch in branches)
            {
                // Selection rule is based on Branch name strings (naming convention dependency).
                if (branch[0].Value.Contains("steamgas") && branch[0].Value.Contains("AHU Main Branch"))
                {
                    UpdateBranch(idfReader, branch);
                }
            }
        }

        private void UpdateBranch(IdfReader reader, IdfObject branch)
        {
            const string gasHumidifierObjectType = "Humidifier:Steam:Gas";
            const string electricHumidifierObjectType = "Humidifier:Steam:Electric";

            // Scan Branch fields to find an equipment "Object Type" field equal to Humidifier:Steam:Electric.
            foreach (int i in Enumerable.Range(1, (branch.Count - 2)))
            {
                Field objectTypeField = branch.Fields[i];
                Field objectNameField = branch.Fields[i + 1];

                if (objectTypeField.Comment.ToLower().Contains("object type") &&
                    objectTypeField.Value.ToLower() == electricHumidifierObjectType.ToLower())
                {
                    IdfObject electricHumidifier = FindObject(reader, electricHumidifierObjectType, objectNameField.Value);

                    // Update Branch reference to point to the gas humidifier object type (name stays the same).
                    objectTypeField.Value = gasHumidifierObjectType;

                    // Create a new Humidifier:Steam:Gas object by copying key fields from the electric humidifier.
                    string humidifierText = GetHumidifierIdfText(electricHumidifier, gasHumidifierObjectType);
                    reader.Load(humidifierText);

                    MessageBox.Show("Replacing electric humidifier: " + electricHumidifier[0] + " with gas humidifier."); // Comment this line to remove message box

                    // Remove the original electric humidifier object from the IDF.
                    reader.Remove(electricHumidifier);
                    break;
                }
            }
        }

        private string GetHumidifierIdfText(IdfObject electricHumidifier, string gasHumidifierObjectType)
        {
            // Electric humidifier fields are used to populate the gas humidifier fields.
            string name = electricHumidifier[0].Value;
            string availability = electricHumidifier[1].Value;
            string ratedCapacity = electricHumidifier[2].Value;
            string ratedGasRate = electricHumidifier[3].Value;
            string thermalEfficiency = "0.9"; // USER CONFIGURATION: Set thermal efficiency (default: 0.9)
            string thermalEfficiencyCurve = "";
            string ratedFanPower = electricHumidifier[4].Value;
            string auxPower = electricHumidifier[5].Value;
            string inletNodeName = electricHumidifier[6].Value;
            string outletNodeName = electricHumidifier[7].Value;

            string[] fields =
            {
                gasHumidifierObjectType,
                name,
                availability,
                ratedCapacity,
                ratedGasRate,
                thermalEfficiency,
                thermalEfficiencyCurve,
                ratedFanPower,
                auxPower,
                inletNodeName,
                outletNodeName
            };

            return String.Join(",", fields) + ";";
        }

        private void AddOutputs(IdfReader reader)
        {
            // USER CONFIGURATION: Add gas-rate reporting for humidifiers (by default, hourly and run-period timestamps).
            reader.Load("Output:Variable,*,Humidifier NaturalGas Rate,hourly;");
            reader.Load("Output:Variable,*,Humidifier NaturalGas Rate,Runperiod;");
        }

        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            UpdateBranches(idfReader);
            AddOutputs(idfReader);

            idfReader.Save();
        }
    }
}