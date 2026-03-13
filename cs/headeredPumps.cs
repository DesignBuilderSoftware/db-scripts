/*
Headered Pump Replacement Script

This DesignBuilder C# script converts standard Pump:VariableSpeed and Pump:ConstantSpeed pump objects into headered pump equivalents (HeaderedPumps:*).

Purpose

For each configured pump:
1) Locate the original Pump:* object by name.
2) Create an equivalent HeaderedPumps:* object by taking fields from the original pump.
3) Update references in Branch objects from Pump:* to HeaderedPumps:* for the same pump name.
4) Insert (Load) the new HeaderedPumps:* object and remove the original Pump:* object.

How to Use

Configuration
- Configure pumps to be replaced (pumpType, pumpName, number of pumps in headered pump bank)
- All pump attributes are taken from DesignBuilder User Interface (HVAC layout).

Prerequisites (required placeholders)
- The base model must already contain the target pump objects:
  - Pump:ConstantSpeed and/or Pump:VariableSpeed with names matching your configuration.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System.Linq;
using System.Collections.Generic;
using System;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        public override void BeforeEnergySimulation()
        {
            IdfReader idf = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // ----------------------------
            // USER CONFIGURATION SECTION
            // ----------------------------
            // Configure pumps to be replaced (pumpType, pumpName, pumps-in-bank)
            ReplacePump(idf, "Pump:ConstantSpeed", "HW Loop Supply Pump", nPumpsInBank: 3);
            ReplacePump(idf, "Pump:VariableSpeed", "CHW Loop Supply Pump", nPumpsInBank: 3);

            idf.Save();
        }

        public IdfObject FindObject(IdfReader idf, string objectType, string objectName)
        {
            try
            {
                return idf[objectType].First(c => c[0] == objectName);
            }
            catch (Exception)
            {
                throw new Exception(String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        public void ReplacePump(IdfReader idf, string pumpType, string pumpName, int nPumpsInBank)
        {
            // Locate the original pump placeholder object
            IdfObject pump = FindObject(idf, pumpType, pumpName);

            // Build the replacement HeaderedPumps:* IDF text from the original pump fields
            string headeredPumpIdfText;
            string pumpTypeLower = pumpType.ToLower();

            if (pumpTypeLower == "pump:variablespeed")
            {
                headeredPumpIdfText = GetVariablePump(pump, nPumpsInBank);

                // Update Branch references from Pump:VariableSpeed -> HeaderedPumps:VariableSpeed (same name)
                ReplaceObjectTypeInList(idf, "Branch", "Pump:VariableSpeed", pumpName, "HeaderedPumps:VariableSpeed", pumpName);
            }
            else if (pumpTypeLower == "pump:constantspeed")
            {
                headeredPumpIdfText = GetConstantPump(pump, nPumpsInBank);

                // Update Branch references from Pump:ConstantSpeed -> HeaderedPumps:ConstantSpeed (same name)
                ReplaceObjectTypeInList(idf, "Branch", "Pump:ConstantSpeed", pumpName, "HeaderedPumps:ConstantSpeed", pumpName);
            }
            else
            {
                throw new Exception(String.Format("Invalid pump type {0}", pumpType));
            }

            // Insert the new headered pump object, then remove the original pump object to avoid duplicates
            idf.Load(headeredPumpIdfText);
            idf.Remove(pump);
        }

        public string GetConstantPump(IdfObject pump, int nPumpsInBank)
        {
            // Maps fields from Pump:ConstantSpeed into HeaderedPumps:ConstantSpeed (by position)
            string headeredPumpTemplate = @"
HeaderedPumps:ConstantSpeed,
  {0},           !- Name
  {1},           !- Inlet Node Name
  {2},           !- Outlet Node Name
  {3},           !- Total Design Flow Rate
  {4},           !- Number of Pumps in Bank
  SEQUENTIAL,    !- Flow Sequencing Control Scheme
  {5},           !- Design Pump Head
  {6},           !- Design Power Consumption
  {7},           !- Motor Efficiency
  {8},           !- Fraction of Motor Inefficiencies to Fluid Stream
  {9};           !- Pump Control Type";

            return String.Format(
                headeredPumpTemplate,
                pump[0].Value, pump[1].Value, pump[2].Value, pump[3].Value, nPumpsInBank.ToString(), pump[4].Value, pump[5].Value, pump[6].Value, pump[7].Value, pump[8].Value);
        }

        public string GetVariablePump(IdfObject pump, int nPumpsInBank)
        {
            // Maps fields from Pump:VariableSpeed into HeaderedPumps:VariableSpeed (by position)
            string headeredPumpTemplate = @"      
HeaderedPumps:VariableSpeed,
  {0},           !- Name
  {1},           !- Inlet Node Name
  {2},           !- Outlet Node Name
  {3},           !- Total Design Flow Rate m3/s
  {4},           !- Number of Pumps in Bank
  SEQUENTIAL,    !- Flow Sequencing Control Scheme
  {5},           !- Design Pump Head Pa
  {6},           !- Design Power Consumption W
  {7},           !- Motor Efficiency
  {8},           !- Fraction of Motor Inefficiencies to Fluid Stream
  {9},           !- Coefficient 1 of the Part Load Performance Curve
  {10},          !- Coefficient 2 of the Part Load Performance Curve
  {11},          !- Coefficient 3 of the Part Load Performance Curve
  {12},          !- Coefficient 4 of the Part Load Performance Curve
  {13},          !- Minimum Flow Rate m3/s
  {14};          !- Pump Control Type";

            return String.Format(
                headeredPumpTemplate,
                pump[0].Value, pump[1].Value, pump[2].Value, pump[3].Value, nPumpsInBank, pump[4].Value, pump[5].Value, pump[6].Value, pump[7].Value, pump[8].Value, pump[9].Value, pump[10].Value, pump[11].Value, pump[12].Value, pump[13].Value);
        }

        private void ReplaceObjectTypeInList(IdfReader idf, string listName, string oldObjectType, string oldObjectName, string newObjectType, string newObjectName)
        {
            // This scans objects of type listName (currently used with "Branch") and replaces (type,name) pairs in fields
            IEnumerable<IdfObject> listObjects = idf[listName];
            bool objectFound = false;

            foreach (IdfObject equipment in listObjects)
            {
                if (!objectFound)
                {
                    for (int i = 0; i < (equipment.Count - 1); i++)
                    {
                        Field field = equipment[i];
                        Field nextField = equipment[i + 1];

                        if (field.Value.ToLower() == oldObjectType.ToLower() &&
                            nextField.Value.ToLower() == oldObjectName.ToLower())
                        {
                            field.Value = newObjectType;
                            nextField.Value = newObjectName;
                            objectFound = true;
                            break;
                        }
                    }
                }
            }
        }
    }
}