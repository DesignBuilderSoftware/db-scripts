/*
Heat Exchanger Conversion Script (Sensible+Latent to Flat Plate)

Purpose:
This DesignBuilder C# script replaces specified EnergyPlus HeatExchanger:AirToAir:SensibleAndLatent objects
with HeatExchanger:AirToAir:FlatPlate objects, while keeping the same object Name.

Main Steps:
1) Locate each target HeatExchanger:AirToAir:SensibleAndLatent object by name
2) Update AirLoopHVAC:OutdoorAirSystem:EquipmentList references for the HeatExchanger:AirToAir:* from SensibleAndLatent to FlatPlate
3) Create a new HeatExchanger:AirToAir:FlatPlate object using a boilerplate IDF template
     (with key attributes copied from the original HX)
4) Remove the original HX object and loading the new FlatPlate object into the IDF

How to Use:

Configuration
- targetHeatExchangerNames:
    List of exact object names to convert (must match the IDF "Name" field exactly).
- flatPlateBoilerplateIdf:
    IDF text template for HeatExchanger:AirToAir:FlatPlate. 
    Placeholders {0}..{8} are populated from the original HX object fields via the user interface.

Prerequisites / Placeholders

Base model must contain, for each target name:
- A HeatExchanger:AirToAir:SensibleAndLatent object with that exact Name
- The original HX object must have values for fields read by this script:
  Availability Schedule Name, Economizer Lockout, Nominal Supply Air Flow Rate, Nominal Electric Power, and the four node names.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. 
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DB.Extensibility.Contracts;
using System;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        // ----------------------------
        // USER CONFIGURATION SECTION
        // ----------------------------
        // List of HeatExchanger:AirToAir:SensibleAndLatent object names to convert
        string[] targetHeatExchangerNames = new string[] { "Air Loop AHU Heat Recovery Device" };

        // ----------------------------
        // USER CONFIGURATION SECTION
        // ----------------------------
        // Boilerplate IDF template for the replacement HeatExchanger:AirToAir:FlatPlate object.
        // Edit fixed attributes here if needed; placeholders {0}..{8} are filled from the original HX.
        string flatPlateBoilerplateIdf = @"
HeatExchanger:AirToAir:FlatPlate,
  {0},                        !- Name
  {1},                        !- Availability Schedule Name
  CounterFlow,                !- Flow Arrangement Type
  {2},                        !- Economizer Lockout
  1,                          !- Ratio of Supply to Secondary hA Values
  {3},                        !- Nominal Supply Air Flow Rate m3/s
  5.0,                        !- Nominal Supply Air Inlet Temperature C
  15.0,                       !- Nominal Supply Air Outlet Temperature C
  {3},                        !- Nominal Secondary Air Flow Rate m3/s
  20.0,                       !- Nominal Secondary Air Inlet Temperature C
  {4},                        !- Nominal Electric Power W
  {5},                        !- Supply Air Inlet Node Name
  {6},                        !- Supply Air Outlet Node Name
  {7},                        !- Secondary Air Inlet Node Name
  {8};                        !- Secondary Air Outlet Node Name";

        private IdfObject FindObject(IdfReader idfReader, string objectType, string objectName)
        {
            return idfReader[objectType].First(o => o[0] == objectName);
        }

        private void ReplaceObjectTypeInList(
          IdfReader idfReader,
          string listName,
          string oldObjectType,
          string oldObjectName,
          string newObjectType,
          string newObjectName)
        {
            // Updates an object type+name pair inside a list object (e.g., EquipmentList)
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

                        if (field.Value == oldObjectType && nextField.Value == oldObjectName)
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

        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
              ApiEnvironment.EnergyPlusInputIdfPath,
              ApiEnvironment.EnergyPlusInputIddPath
              );

            string oldObjectType = "HeatExchanger:AirToAir:SensibleAndLatent";
            string newObjectType = "HeatExchanger:AirToAir:FlatPlate";

            foreach (string objectName in targetHeatExchangerNames)
            {
                // 1) Find the existing Sensible+Latent HX (must exist as a placeholder)
                IdfObject oldHx = FindObject(idfReader, oldObjectType, objectName);

                // 2) Update the Outdoor Air System equipment list to reference FlatPlate instead
                ReplaceObjectTypeInList(
                  idfReader,
                  "AirLoopHVAC:OutdoorAirSystem:EquipmentList",
                  oldObjectType,
                  objectName,
                  newObjectType,
                  objectName);

                // 3) Copy required fields from old HX to populate the FlatPlate template
                string name = oldHx["Name"].Value;
                string availabilityScheduleName = oldHx["Availability Schedule Name"].Value;
                string economizerLockout = oldHx["Economizer Lockout"].Value;
                string nominalSupplyAirFlowRate = oldHx["Nominal Supply Air Flow Rate"].Value;
                string nominalElectricPower = oldHx["Nominal Electric Power"].Value;

                string supplyInletNode = oldHx["Supply Air Inlet Node Name"].Value;
                string supplyOutletNode = oldHx["Supply Air Outlet Node Name"].Value;
                string exhaustInletNode = oldHx["Exhaust Air Inlet Node Name"].Value;
                string exhaustOutletNode = oldHx["Exhaust Air Outlet Node Name"].Value;

                // 4) Create FlatPlate IDF text and replace the object in the IDF
                string newFlatPlateIdf = String.Format(
                  flatPlateBoilerplateIdf,
                  name,
                  availabilityScheduleName,
                  economizerLockout,
                  nominalSupplyAirFlowRate,
                  nominalElectricPower,
                  supplyInletNode,
                  supplyOutletNode,
                  exhaustInletNode,
                  exhaustOutletNode);

                idfReader.Remove(oldHx);
                idfReader.Load(newFlatPlateIdf);
            }

            idfReader.Save();
        }
    }
}