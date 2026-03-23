/*
Direct-Fired Absorption Chiller-Heater Replacement Script (District Heating/Cooling)

Purpose
This DesignBuilder C# script replaces DistrictHeating and DistrictCooling plant components with a ChillerHeater:Absorption:DirectFired object.

Main Steps:
1) Replace references to DistrictHeating and DistrictCooling in Branch and PlantEquipmentList
   so they point to ChillerHeater:Absorption:DirectFired with the specified chiller name
2) Read node names from the DistrictHeating and DistrictCooling placeholder objects
3) Insert a boilerplate ChillerHeater:Absorption:DirectFired object (plus required curves and OA node list)
4) Remove the original DistrictHeating and DistrictCooling placeholder objects from the IDF
5) Save the modified IDF

How to Use:

Configuration (edit values in BeforeEnergySimulation)

- chillerName: name assigned to the new ChillerHeater:Absorption:DirectFired object
- districtHeatingName: placeholder DistrictHeating object name to be replaced
- districtCoolingName: placeholder DistrictCooling object name to be replaced
- hwLoopName / chwLoopName: currently not used by the script logic (kept as reference)

ChillerHeater properties (edit boilerplate in GetChillerHeaterIdfObjects)
- Adjust performance ratios, temperatures, curves, and any autosizing fields inside the IDF template string.

Prerequisites / Placeholders
The base model must contain:
- A DistrictHeating object named exactly as districtHeatingName (default: "District Heating")
- A DistrictCooling object named exactly as districtCoolingName (default: "District Cooling")

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. 
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Linq;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        private IdfReader idfReader;

        public override void BeforeEnergySimulation()
        {
            idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // ----------------------------
            // USER CONFIGURATION SECTION
            // ----------------------------
            // Ensure these names match objects in the model/IDF
            const string chillerName = "Big Chiller";
            const string hwLoopName = "HW Loop";     // Not used in current logic
            const string chwLoopName = "CHW Loop";   // Not used in current logic
            const string districtHeatingName = "District Heating";
            const string districtCoolingName = "District Cooling";

            ApplyDirectFiredChiller(
                chillerName,
                hwLoopName,
                chwLoopName,
                districtHeatingName,
                districtCoolingName);

            idfReader.Save();
        }

        public void ApplyDirectFiredChiller(
            string chillerName,
            string hwLoopName,
            string chwLoopName,
            string districtHeatingName,
            string districtCoolingName)
        {
            const string chillerType = "ChillerHeater:Absorption:DirectFired";
            const string heatingPlaceholderObjectType = "DistrictHeating";
            const string coolingPlaceholderObjectType = "DistrictCooling";

            // Replace all Branch and PlantEquipmentList references so the plant now points to the new chiller-heater.
            ReplaceObjectTypeInList("Branch", heatingPlaceholderObjectType, districtHeatingName, chillerType, chillerName);
            ReplaceObjectTypeInList("PlantEquipmentList", heatingPlaceholderObjectType, districtHeatingName, chillerType, chillerName);

            ReplaceObjectTypeInList("Branch", coolingPlaceholderObjectType, districtCoolingName, chillerType, chillerName);
            ReplaceObjectTypeInList("PlantEquipmentList", coolingPlaceholderObjectType, districtCoolingName, chillerType, chillerName);

            // Read the hot and chilled water nodes from the DistrictHeating and DistrictCooling placeholders, respectively.
            IdfObject districtHeating = FindObject(heatingPlaceholderObjectType, districtHeatingName);
            string hwInletNode = districtHeating["Hot Water Inlet Node Name"].Value;
            string hwOutletNode = districtHeating["Hot Water Outlet Node Name"].Value;

            IdfObject districtCooling = FindObject(coolingPlaceholderObjectType, districtCoolingName);
            string chwInletNode = districtCooling["Chilled Water Inlet Node Name"].Value;
            string chwOutletNode = districtCooling["Chilled Water Outlet Node Name"].Value;

            // Inject the new ChillerHeater:Absorption:DirectFired object (plus curves and OA node list).
            string chillerHeaterIdf = GetChillerHeaterIdfObjects(chillerName, hwInletNode, hwOutletNode, chwInletNode, chwOutletNode);
            idfReader.Load(chillerHeaterIdf);

            // Remove the placeholders after the new object has been created and referenced.
            idfReader.Remove(districtHeating);
            idfReader.Remove(districtCooling);
        }

        public string GetChillerHeaterIdfObjects(
            string chillerName,
            string hwInletNode,
            string hwOutletNode,
            string chwInletNode,
            string chwOutletNode)
        {
            // ----------------------------
            // USER CONFIGURATION SECTION
            // ----------------------------
            // Boilerplate IDF objects inserted into the model (ChillerHeater:Absorption:DirectFired, OutdoorAir:Nodelist, Performance curves)
            string chillerHeaterTemplate = @"
ChillerHeater:Absorption:DirectFired,
    {0},                      !- Name
    Autosize,                 !- Nominal Cooling Capacity W
    0.8,                      !- Heating to Cooling Capacity Ratio
    0.97,                     !- Fuel Input to Cooling Output Ratio
    1.25,                     !- Fuel Input to Heating Output Ratio
    0.01,                     !- Electric Input to Cooling Output Ratio
    0.005,                    !- Electric Input to Heating Output Ratio
    {3},                      !- Chilled Water Inlet Node Name
    {4},                      !- Chilled Water Outlet Node Name
    {0} Chiller OA Node,      !- Condenser Inlet Node Name
    ,                         !- Condenser Outlet Node Name
    {1},                      !- Hot Water Inlet Node Name
    {2},                      !- Hot Water Outlet Node Name
    0.000001,                 !- Minimum Part Load Ratio
    1.0,                      !- Maximum Part Load Ratio
    0.6,                      !- Optimum Part Load Ratio
    29,                       !- Design Entering Condenser Water Temperature C
    7,                        !- Design Leaving Chilled Water Temperature C
    Autosize,                 !- Design Chilled Water Flow Rate m3/s
    Autosize,                 !- Design Condenser Water Flow Rate m3/s
    Autosize,                 !- Design Hot Water Flow Rate m3/s
    {0} GasAbsFlatBiQuad,     !- Cooling Capacity Function of Temperature Curve Name
    {0} GasAbsFlatBiQuad,     !- Fuel Input to Cooling Output Ratio Function of Temperature Curve Name
    {0} GasAbsLinearQuad,     !- Fuel Input to Cooling Output Ratio Function of Part Load Ratio Curve Name
    {0} GasAbsFlatBiQuad,     !- Electric Input to Cooling Output Ratio Function of Temperature Curve Name
    {0} GasAbsFlatQuad,       !- Electric Input to Cooling Output Ratio Function of Part Load Ratio Curve Name
    {0} GasAbsInvLinearQuad,  !- Heating Capacity Function of Cooling Capacity Curve Name
    {0} GasAbsLinearQuad,     !- Fuel Input to Heat Output Ratio During Heating Only Operation Curve Name
    EnteringCondenser,        !- Temperature Curve Input Variable
    AirCooled,                !- Condenser Type
    2,                        !- Chilled Water Temperature Lower Limit C
    0,                        !- Fuel Higher Heating Value kJ/kg
    NaturalGas,               !- Fuel Type
    ;                         !- Sizing Factor

OutdoorAir:Nodelist,
    {0} Chiller OA Node;      !- Outside air node

Curve:Biquadratic,
    {0} GasAbsFlatBiQuad,     !- Name
    1.000000000,             !- Coefficient1 Constant
    0.000000000,             !- Coefficient2 x
    0.000000000,             !- Coefficient3 x**2
    0.000000000,             !- Coefficient4 y
    0.000000000,             !- Coefficient5 y**2
    0.000000000,             !- Coefficient6 x*y
    0.,                      !- Minimum Value of x
    50.,                     !- Maximum Value of x
    0.,                      !- Minimum Value of y
    50.;                     !- Maximum Value of y

Curve:Quadratic,
    {0} GasAbsFlatQuad,       !- Name
    1.000000000,             !- Coefficient1 Constant
    0.000000000,             !- Coefficient2 x
    0.000000000,             !- Coefficient3 x**2
    0.,                      !- Minimum Value of x
    50.;                     !- Maximum Value of x

Curve:Quadratic,
    {0} GasAbsLinearQuad,     !- Name
    0.000000000,             !- Coefficient1 Constant
    1.000000000,             !- Coefficient2 x
    0.000000000,             !- Coefficient3 x**2
    0.,                      !- Minimum Value of x
    50.;                     !- Maximum Value of x

Curve:Quadratic,
    {0} GasAbsInvLinearQuad,  !- Name
    1.000000000,             !- Coefficient1 Constant
    -1.000000000,            !- Coefficient2 x
    0.000000000,             !- Coefficient3 x**2
    0.,                      !- Minimum Value of x
    50.;                     !- Maximum Value of x
";

            return string.Format(
                chillerHeaterTemplate,
                chillerName,
                hwInletNode,
                hwOutletNode,
                chwInletNode,
                chwOutletNode);
        }

        public IdfObject FindObject(string objectType, string objectName)
        {
            try
            {
                return idfReader[objectType].First(c => c[0] == objectName);
            }
            catch (Exception)
            {
                throw new MissingFieldException(string.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        private void ReplaceObjectTypeInList(
            string listName,
            string oldObjectType,
            string oldObjectName,
            string newObjectType,
            string newObjectName)
        {
            var idfObjects = idfReader[listName];

            bool replacementMade = false;

            foreach (IdfObject idfObject in idfObjects)
            {
                if (replacementMade)
                {
                    break;
                }

                for (int i = 0; i < (idfObject.Count - 1); i++)
                {
                    Field currentField = idfObject[i];
                    Field nextField = idfObject[i + 1];

                    // Note: comparison is case-sensitive here.
                    if (currentField.Value == oldObjectType && nextField.Value == oldObjectName)
                    {
                        currentField.Value = newObjectType;
                        nextField.Value = newObjectName;
                        replacementMade = true;
                        break;
                    }
                }
            }
        }
    }
}