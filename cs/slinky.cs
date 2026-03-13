/*
Replace Ground Heat Exchanger of type Surface with Slinky type.

Purpose
This DesignBuilder C# script replaces a single GroundHeatExchanger:Surface object with a
GroundHeatExchanger:Slinky object, keeping the same object name and reusing the original inlet/outlet nodes.

Main steps
1) Find the target GroundHeatExchanger:Surface by name
2) Update references in:
   - CondenserEquipmentList (object type + name pair)
   - Branch (object type + name pair)
3) Add:
   - GroundHeatExchanger:Slinky (generated from the boilerplate template)
   - Site:GroundTemperature:Undisturbed:KusudaAchenbach (boilerplate object)
4) Remove the original GroundHeatExchanger:Surface
5) Save the modified IDF

How to Use

Configuration
- Set the target object name in: groundHxName
  This must match the Name field of the GroundHeatExchanger:Surface in the model.
- Adjust the Slinky parameters in: slinkyBoilerplateIdf
  (e.g., design flow rate, soil properties, trench geometry, etc.)
- Adjust / replace the undisturbed ground temperature object in: undisturbedGroundTempsIdf
  Ensure the object name matches the reference used by the slinky boilerplate.

Prerequisites (required placeholders)
Base model must contain a GroundHeatExchanger:Surface object (referenced in groundHxName)

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
        // USER CONFIGURATION:name of GroundHeatExchanger:Surface object (must exactly match the IDF object's Name).
        private string groundHxName = "Ground Heat Exchanger";

        // ----------------------------
        // USER CONFIGURATION SECTION
        // ----------------------------
        // Boilerplate IDF template for the replacement GroundHeatExchanger:Slinky. Edit attribute values here as needed.
        private string slinkyBoilerplateIdf = @"
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

        // ----------------------------
        // USER CONFIGURATION SECTION
        // ----------------------------
        // Boilerplate IDF for undisturbed ground temperature object referenced by the slinky HX above.
        // Ensure the object name matches the reference used in slinkyBoilerplateIdf (e.g., "KATemps").
        private string undisturbedGroundTempsIdf = @"
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

        private void ReplaceObjectTypeInList(
            IdfReader idfReader,
            string listName,
            string oldObjectType,
            string oldObjectName,
            string newObjectType,
            string newObjectName)
        {
            // Updates (Object Type, Object Name) pairs inside list-like objects.
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

            // Required placeholder object: GroundHeatExchanger:Surface with Name == groundHxName
            IdfObject surfaceGroundHx = FindObject(idfReader, oldObjectType, groundHxName);

            // Update reference locations where the ground HX is referenced by (type, name).
            ReplaceObjectTypeInList(idfReader, "CondenserEquipmentList", oldObjectType, groundHxName, newObjectType, groundHxName);
            ReplaceObjectTypeInList(idfReader, "Branch", oldObjectType, groundHxName, newObjectType, groundHxName);

            // Reuse inlet/outlet node names from the original surface HX so connectivity remains consistent.
            string inletNode = surfaceGroundHx["Fluid Inlet Node Name"].Value;
            string outletNode = surfaceGroundHx["Fluid Outlet Node Name"].Value;

            // Create the replacement slinky object using the boilerplate template.
            string slinkyGroundHxIdf = String.Format(slinkyBoilerplateIdf, groundHxName, inletNode, outletNode);

            // Replace objects in the IDF: remove old, add new + supporting ground temperature object.
            idfReader.Remove(surfaceGroundHx);
            idfReader.Load(slinkyGroundHxIdf);
            idfReader.Load(undisturbedGroundTempsIdf);

            idfReader.Save();
        }
    }
}