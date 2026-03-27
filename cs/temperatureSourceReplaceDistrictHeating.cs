/*
Replace District Heating with Plant Temperature Source component

Purpose:
This DesignBuilder C# script replaces "DistrictHeating" with "PlantComponent:TemperatureSource".
This object used to simulate systems with a known supply temperature (e.g., rivers, wells, and other configurations where a known temperature is pumped back into the plant system).

Main Steps:
1) Finds a placeholder DistrictHeating object by name
2) Replaces references to that DistrictHeating object in PlantEquipmentList and Branch
3) Creates a PlantComponent:TemperatureSource object using:
   - The same object name as the placeholder DistrictHeating object
   - The inlet/outlet nodes taken from the DistrictHeating object
   - A constant source temperature defined in this script
4) Removes the original DistrictHeating object and saves the updated IDF

How to Use:

Configuration
- Set 'objectName' to the Name of the DistrictHeating placeholder you want to convert.
- Set 'sourceTemperatureC' to the desired constant temperature in °C.

Prerequisites / Placeholders
- The base model must contain a DistrictHeating placeholder object.

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
        // IDF template for the new PlantComponent:TemperatureSource object.
        private readonly string temperatureSourceIdfTemplate = @"
PlantComponent:TemperatureSource,
      {0},                     !- Name
      {1},                     !- Inlet Node
      {2},                     !- Outlet Node
      Autosize,                !- Design Volume Flow Rate m3/s
      Constant,                !- Temperature Specification Type
      {3},                     !- Source Temperature C
      ;                        !- Source Temperature Schedule Name";

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
            IEnumerable<IdfObject> allEquipment = idfReader[listName];

            foreach (IdfObject equipment in allEquipment)
            {
                for (int i = 0; i < (equipment.Count - 1); i++)
                {
                    Field field = equipment[i];
                    Field nextField = equipment[i + 1];

                    if (field.Value == oldObjectType && nextField.Value == oldObjectName)
                    {
                        field.Value = newObjectType;
                        nextField.Value = newObjectName;
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

            // USER CONFIGURATION: Name of the placeholder DistrictHeating object to replace
            string objectName = "TemperatureSource";
            string oldObjectType = "DistrictHeating";
            string newObjectType = "PlantComponent:TemperatureSource";

            // USER CONFIGURATION: Constant temperature to apply for PlantComponent:TemperatureSource
            int sourceTemperatureC = 10;

            // Find the DistrictHeating placeholder object that provides inlet/outlet nodes
            IdfObject districtHeating = FindObject(idfReader, oldObjectType, objectName);

            ReplaceObjectTypeInList(idfReader, "PlantEquipmentList", oldObjectType, objectName, newObjectType, objectName);
            ReplaceObjectTypeInList(idfReader, "Branch", oldObjectType, objectName, newObjectType, objectName);

            string inletNode = districtHeating["Hot Water Inlet Node Name"].Value;
            string outletNode = districtHeating["Hot Water Outlet Node Name"].Value;

            // Build the replacement object IDF text
            string temperatureSource = String.Format(
                temperatureSourceIdfTemplate,
                objectName,
                inletNode,
                outletNode,
                sourceTemperatureC);

            // Replace the placeholder object with the new component and save the modified IDF
            idfReader.Remove(districtHeating);
            idfReader.Load(temperatureSource);

            idfReader.Save();
        }
    }
}