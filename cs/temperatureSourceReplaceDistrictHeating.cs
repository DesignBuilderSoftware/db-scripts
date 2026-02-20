/*
Replace "DistrictHeating" with "PlantComponent:TemperatureSource".

Object name needs to reference the DistrictHeating object name.
Attributes can be set in 'boilerplate' IDF text below.
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

        string boilerplate = @"
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

        private void ReplaceObjectTypeInList(IdfReader idfReader, string listName, string oldObjectType, string oldObjectName, string newObjectType, string newObjectName)
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
            string objectName = "TemperatureSource";
            string oldObjectType = "DistrictHeating";
            string newObjectType = "PlantComponent:TemperatureSource";
            int temperature = 10;

            IdfObject districtHeating = FindObject(idfReader, oldObjectType, objectName);

            ReplaceObjectTypeInList(idfReader, "PlantEquipmentList", oldObjectType, objectName, newObjectType, objectName);
            ReplaceObjectTypeInList(idfReader, "Branch", oldObjectType, objectName, newObjectType, objectName);

            string inletNode = districtHeating["Hot Water Inlet Node Name"].Value;
            string outletNode = districtHeating["Hot Water Outlet Node Name"].Value;

            string temperatureSource = String.Format(boilerplate, objectName, inletNode, outletNode, temperature);

            idfReader.Remove(districtHeating);
            idfReader.Load(temperatureSource);

            idfReader.Save();
        }
    }
}