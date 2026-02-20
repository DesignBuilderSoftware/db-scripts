/*
Replace GroundHeatExchanger:Surface with GroundHeatExchanger:Slinky.

Object name needs to reference the GroundHeatExchanger:Surface object name.
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
        string objectName = "Ground Heat Exchanger";
        string slinkyBoilerPlate = @"
GroundHeatExchanger:Slinky,
  {0},              !- Name
  {1},              !- Inlet Node
  {2},              !- Outlet Node
  0.0033,           !- Design Flow Rate [m3/s]
  1.2,              !- Soil Thermal Conductivity [W/m-K]
  3200,             !- Soil Density [kg/m3]
  850,              !- Soil Specific Heat [J/kg-K]
  1.8,              !- Pipe Thermal Conductivity [W/m-K]
  920,              !- Pipe Density [kg/m3]
  2200,             !- Pipe Specific Heat [J/kg-K]
  0.02667,          !- Pipe Outside Diameter [m]
  0.002413,         !- Pipe Wall Thickness [m]
  Horizontal,       !- Heat Exchanger Configuration (Vertical, Horizontal)
  1,                !- Coil Diameter [m]
  0.2,              !- Coil Pitch [m]
  2.5,              !- Trench Depth [m]
  40,               !- Trench Length [m]
  15,               !- Number of Parallel Trenches
  2,                !- Trench Spacing [m]
  Site:GroundTemperature:Undisturbed:KusudaAchenbach, !- Type of Undisturbed Ground Temperature Object
  KATemps,          !- Name of Undisturbed Ground Temperature Object
  10;               !- Maximum length of simulation [years]";

        string ground = @"
Site:GroundTemperature:Undisturbed:KusudaAchenbach,
  KATemps,                 !- Name
  1.8,                     !- Soil Thermal Conductivity {W/m-K}
  920,                     !- Soil Density {kg/m3}
  2200,                    !- Soil Specific Heat {J/kg-K}
  15.5,                    !- Average Soil Surface Temperature {C}
  3.2,                     !- Average Amplitude of Surface Temperature {deltaC}
  8;                       !- Phase Shift of Minimum Surface Temperature {days}";

        private IdfObject FindObject(IdfReader idfReader, string objectType, string objectName)
        {
            return idfReader[objectType].First(o => o[0] == objectName);
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

            string oldObjectType = "GroundHeatExchanger:Surface";
            string newObjectType = "GroundHeatExchanger:Slinky";

            IdfObject groundHX = FindObject(idfReader, oldObjectType, objectName);

            ReplaceObjectTypeInList(idfReader, "CondenserEquipmentList", oldObjectType, objectName, newObjectType, objectName);
            ReplaceObjectTypeInList(idfReader, "Branch", oldObjectType, objectName, newObjectType, objectName);

            string inletNode = groundHX["Fluid Inlet Node Name"].Value;
            string outletNode = groundHX["Fluid Outlet Node Name"].Value;

            string slinky = String.Format(slinkyBoilerPlate, objectName, inletNode, outletNode);

            idfReader.Remove(groundHX);
            idfReader.Load(slinky);
            idfReader.Load(ground);

            idfReader.Save();
        }
    }
}





