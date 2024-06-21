/*
This script replaces Pump:VariableSpeed and Pump:ConstantSpeed objects with their
headered equivalents.

All pump attributes are cherry-picked from DesignBuilder HVAC layout.

*/
using System.Runtime;
using System.Linq;
using System.Collections.Generic;
using System;
using System.Windows.Forms;
using DB.Extensibility.Contracts;
using EpNet;


namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // define pumps to be replaced
            ReplacePump(idfReader, "Pump:ConstantSpeed", "HW Loop Supply Pump", nPumpsInBank: 3);
            ReplacePump(idfReader, "Pump:VariableSpeed", "CHW Loop Supply Pump", nPumpsInBank: 3);

            idfReader.Save();
        }
        public IdfObject FindObject(IdfReader reader, string objectType, string objectName)
        {
            try
            {
                return reader[objectType].First(c => c[0] == objectName);
            }
            catch(Exception e)
            {
                throw new Exception(String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        public void ReplacePump(IdfReader reader, string pumpType, string pumpName, int nPumpsInBank)
        {
            IdfObject pump = FindObject(reader, pumpType, pumpName);
            string headeredPump;

            if (pumpType.ToLower() == "pump:variablespeed")
            {
                headeredPump = GetVariablePump(pump, nPumpsInBank);
                ReplaceObjectTypeInList(reader, "Branch", "Pump:VariableSpeed", pumpName, "HeaderedPumps:VariableSpeed", pumpName);
            }
            else if (pumpType.ToLower() == "pump:constantspeed")
            {
                headeredPump = GetConstantPump(pump, nPumpsInBank);
                ReplaceObjectTypeInList(reader, "Branch", "Pump:ConstantSpeed", pumpName, "HeaderedPumps:ConstantSpeed", pumpName);
            } else {
                throw new Exception(String.Format("Invalid pump type {0}", pumpType));
            }
            reader.Load(headeredPump);
            reader.Remove(pump);
        }

        public string GetConstantPump(IdfObject pump, int nPumpsInBank)
        {
        string headeredPump = @"HeaderedPumps:ConstantSpeed,
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
         return String.Format(headeredPump, pump[0].Value, pump[1].Value, pump[2].Value, pump[3].Value, nPumpsInBank.ToString(), pump[4].Value, pump[5].Value, pump[6].Value, pump[7].Value, pump[8].Value);
        }

        public string GetVariablePump(IdfObject pump, int nPumpsInBank)
        {
        string headeredPump = @"HeaderedPumps:VariableSpeed,
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


         return String.Format(headeredPump, pump[0].Value, pump[1].Value, pump[2].Value, pump[3].Value, nPumpsInBank, pump[4].Value, pump[5].Value, pump[6].Value, pump[7].Value, pump[8].Value, pump[9].Value, pump[10].Value, pump[11].Value, pump[12].Value, pump[13].Value);
        }

        private void ReplaceObjectTypeInList(IdfReader idfReader, string listName, string oldObjectType, string oldObjectName, string newObjectType, string newObjectName)
        {
            IEnumerable<IdfObject> allEquipment = idfReader[listName];

            bool objectFound = false;

            foreach (IdfObject equipment in allEquipment)
            {
                if (!objectFound)
                {
                    for (int i = 0; i < (equipment.Count - 1); i++)
                    {
                        Field field = equipment[i];
                        Field nextField = equipment[i + 1];

                        if (field.Value.ToLower() == oldObjectType.ToLower() && nextField.Value.ToLower() == oldObjectName.ToLower())
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
