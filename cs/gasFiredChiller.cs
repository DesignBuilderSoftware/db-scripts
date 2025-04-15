/*
Replace District Heating and District Cooling source plant components with ChillerHeater:Absorption:DirectFired object.

ChillerHeater properties can be adjusted in the string boilerplate.

 */

using System.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        private IdfReader Reader;

        public override void BeforeEnergySimulation()
        {
            Reader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            const string chillerName = "Big Chiller";
            const string hwLoopName = "HW Loop";
            const string chwLoopName = "CHW Loop";
            const string districtHeatingName = "District Heating";
            const string districtCoolingName = "District Cooling";

            ApplyDirectFiredChiller(
                chillerName, 
                hwLoopName, 
                chwLoopName, 
                districtHeatingName,
                districtCoolingName);

            Reader.Save();
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

            // update placeholder list and branch references
            ReplaceObjectTypeInList("Branch", heatingPlaceholderObjectType, districtHeatingName, chillerType,
                chillerName);
            ReplaceObjectTypeInList("PlantEquipmentList", heatingPlaceholderObjectType, districtHeatingName,
                chillerType, chillerName);

            ReplaceObjectTypeInList("Branch", coolingPlaceholderObjectType, districtCoolingName, chillerType,
                chillerName);
            ReplaceObjectTypeInList("PlantEquipmentList", coolingPlaceholderObjectType, districtCoolingName,
                chillerType, chillerName);

            // modify nodes
            IdfObject districtHeating = FindObject(heatingPlaceholderObjectType, districtHeatingName);
            string hwInletNode = districtHeating["Hot Water Inlet Node Name"].Value;
            string hwOutletNode = districtHeating["Hot Water Outlet Node Name"].Value;

            IdfObject districtCooling = FindObject(coolingPlaceholderObjectType, districtCoolingName);
            string chwInletNode = districtCooling["Chilled Water Inlet Node Name"].Value;
            string chwOutletNode = districtCooling["Chilled Water Outlet Node Name"].Value;

            string chillerHeater = GetChillerHeaterIdfObjects(chillerName, hwInletNode, hwOutletNode, chwInletNode, chwOutletNode);
            Reader.Load(chillerHeater);

            Reader.Remove(districtHeating);
            Reader.Remove(districtCooling);
        }

        public string GetChillerHeaterIdfObjects(string chillerName, string hwInletNode,
            string hwOutletNode, string chwInletNode, string chwOutletNode)
        {
            string template = @"ChillerHeater:Absorption:DirectFired,
    {0},                      !- Name
    Autosize,                  !- Nominal Cooling Capacity W
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


 OutdoorAir:Nodelist, {0} Chiller OA Node;   ! - Outside air node

 Curve:Biquadratic,
    {0} GasAbsFlatBiQuad,     !- Name
    1.000000000,              !- Coefficient1 Constant
    0.000000000,              !- Coefficient2 x
    0.000000000,              !- Coefficient3 x**2
    0.000000000,              !- Coefficient4 y
    0.000000000,              !- Coefficient5 y**2
    0.000000000,              !- Coefficient6 x*y
    0.,                       !- Minimum Value of x
    50.,                      !- Maximum Value of x
    0.,                       !- Minimum Value of y
    50.;                      !- Maximum Value of y

  Curve:Quadratic,
    {0} GasAbsFlatQuad,       !- Name
    1.000000000,              !- Coefficient1 Constant
    0.000000000,              !- Coefficient2 x
    0.000000000,              !- Coefficient3 x**2
    0.,                       !- Minimum Value of x
    50.;                      !- Maximum Value of x

  Curve:Quadratic,
    {0} GasAbsLinearQuad,     !- Name
    0.000000000,              !- Coefficient1 Constant
    1.000000000,              !- Coefficient2 x
    0.000000000,              !- Coefficient3 x**2
    0.,                       !- Minimum Value of x
    50.;                      !- Maximum Value of x

  Curve:Quadratic,
    {0} GasAbsInvLinearQuad,  !- Name
    1.000000000,              !- Coefficient1 Constant
    -1.000000000,             !- Coefficient2 x
    0.000000000,              !- Coefficient3 x**2
    0.,                       !- Minimum Value of x
    50.;                      !- Maximum Value of x

";
            return String.Format(template, chillerName, hwInletNode, hwOutletNode, chwInletNode,
                chwOutletNode);
        }

        public IdfObject FindObject(string objectType, string objectName)
        {
            try
            {
                return Reader[objectType].First(c => c[0] == objectName);
            }
            catch (Exception e)
            {
                throw new MissingFieldException(String.Format("Cannot find object: {0}, type: {1}", objectName,
                    objectType));
            }
        }

        private void ReplaceObjectTypeInList(string listName, string oldObjectType, string oldObjectName,
            string newObjectType, string newObjectName)
        {
            IEnumerable<IdfObject> allEquipment = Reader[listName];

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
    }
}