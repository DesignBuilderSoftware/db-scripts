/*
Add Desiccant Dehumidifier (of type Balanced Flow) to an AHU branch

This DesignBuilder C# script inserts an EnergyPlus HeatExchanger:Desiccant:BalancedFlow component into an existing air loop supply branch.

Purpose

1) Locate the AHU main supply Branch for a given air loop
2) Determine the process-side inlet node from the branch field list at the chosen insertion point
3) Create a new process outlet node name and rewire the downstream component to use it
4) Insert a HeatExchanger:Desiccant:BalancedFlow component into the Branch (type/name/inlet/outlet)
5) Add the HeatExchanger:Desiccant:BalancedFlow object plus its PerformanceDataType1 object to the IDF
6) Save the modified IDF

How to Use

Configuration
- insertAtBranchFieldIndex refers to the Branch field list ordering. In a Branch, components are stored in repeated quads:
    (Component 1 Object Type, Component 1 Name, Component 1 Inlet Node, Component 1 Outlet Node,
     Component 2 Object Type, Component 2 Name, Component 2 Inlet Node, Component 2 Outlet Node, ...)
  This script:
   - Uses (insertAtBranchFieldIndex - 1) to read the process inlet node from the existing branch fields
   - Inserts a new quad at insertAtBranchFieldIndex
   - Updates downstream component inlet node and branch outlet-node field using fixed offsets

Prerequisites (required placeholders)

This script expects an Air Loop to be in place in the model, listed in the USER CONFIGURATION SECTION.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Linq;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        // IDF text template: creates the desiccant HX object and a matching performance object (HXDesPerf1).
        private readonly string desiccantHxAndPerfIdfTemplate = @"
HeatExchanger:Desiccant:BalancedFlow,
    {0},                                                       !- Name
    ON 24/7,                                                   !- Availability Schedule Name
    {1},                                                       !- Regeneration Air Inlet Node Name
    {2},                                                       !- Regeneration Air Outlet Node Name
    {3},                                                       !- Process Air Inlet Node Name
    {4},                                                       !- Process Air Outlet Node Name
    HeatExchanger:Desiccant:BalancedFlow:PerformanceDataType1, !- Heat Exchanger Performance Object Type
    HXDesPerf1;                                                !- Heat Exchanger Performance Name

HeatExchanger:Desiccant:BalancedFlow:PerformanceDataType1,
    HXDesPerf1,              !- Name
    1.90,                    !- Nominal Air Flow Rate m3/s
    3.54,                    !- Nominal Air Face Velocity m/s
    50.0,                    !- Nominal Electric Power W
    -7.18302E+00,            !- Temperature Equation Coefficient 1
    -1.84967E+02,            !- Temperature Equation Coefficient 2
    1.00051E+00,             !- Temperature Equation Coefficient 3
    1.16033E+04,             !- Temperature Equation Coefficient 4
    -5.07550E+01,            !- Temperature Equation Coefficient 5
    -1.68467E-02,            !- Temperature Equation Coefficient 6
    5.82213E+01,             !- Temperature Equation Coefficient 7
    5.98863E-01,             !- Temperature Equation Coefficient 8
    0.005143,                !- Minimum Regeneration Inlet Air Humidity Ratio for Temperature Equation kgWater/kgDryAir
    0.024286,                !- Maximum Regeneration Inlet Air Humidity Ratio for Temperature Equation kgWater/kgDryAir
    17.83333,                !- Minimum Regeneration Inlet Air Temperature for Temperature Equation C
    48.88889,                !- Maximum Regeneration Inlet Air Temperature for Temperature Equation C
    0.005000,                !- Minimum Process Inlet Air Humidity Ratio for Temperature Equation kgWater/kgDryAir
    0.017514,                !- Maximum Process Inlet Air Humidity Ratio for Temperature Equation kgWater/kgDryAir
    4.583333,                !- Minimum Process Inlet Air Temperature for Temperature Equation C
    21.83333,                !- Maximum Process Inlet Air Temperature for Temperature Equation C
    2.286,                   !- Minimum Regeneration Air Velocity for Temperature Equation m/s
    4.826,                   !- Maximum Regeneration Air Velocity for Temperature Equation m/s
    16.66667,                !- Minimum Regeneration Outlet Air Temperature for Temperature Equation C
    46.11111,                !- Maximum Regeneration Outlet Air Temperature for Temperature Equation C
    10.0,                    !- Minimum Regeneration Inlet Air Relative Humidity for Temperature Equation percent
    100.0,                   !- Maximum Regeneration Inlet Air Relative Humidity for Temperature Equation percent
    70.0,                    !- Minimum Process Inlet Air Relative Humidity for Temperature Equation percent
    100.0,                   !- Maximum Process Inlet Air Relative Humidity for Temperature Equation percent
    3.13878E-03,             !- Humidity Ratio Equation Coefficient 1
    1.09689E+00,             !- Humidity Ratio Equation Coefficient 2
    -2.63341E-05,            !- Humidity Ratio Equation Coefficient 3
    -6.33885E+00,            !- Humidity Ratio Equation Coefficient 4
    9.38196E-03,             !- Humidity Ratio Equation Coefficient 5
    5.21186E-05,             !- Humidity Ratio Equation Coefficient 6
    6.70354E-02,             !- Humidity Ratio Equation Coefficient 7
    -1.60823E-04,            !- Humidity Ratio Equation Coefficient 8
    0.005143,                !- Minimum Regeneration Inlet Air Humidity Ratio for Humidity Ratio Equation kgWater/kgDryAir
    0.024286,                !- Maximum Regeneration Inlet Air Humidity Ratio for Humidity Ratio Equation kgWater/kgDryAir
    17.83333,                !- Minimum Regeneration Inlet Air Temperature for Humidity Ratio Equation C
    48.88889,                !- Maximum Regeneration Inlet Air Temperature for Humidity Ratio Equation C
    0.005000,                !- Minimum Process Inlet Air Humidity Ratio for Humidity Ratio Equation kgWater/kgDryAir
    0.017514,                !- Maximum Process Inlet Air Humidity Ratio for Humidity Ratio Equation kgWater/kgDryAir
    4.583333,                !- Minimum Process Inlet Air Temperature for Humidity Ratio Equation C
    21.83333,                !- Maximum Process Inlet Air Temperature for Humidity Ratio Equation C
    2.286,                   !- Minimum Regeneration Air Velocity for Humidity Ratio Equation m/s
    4.826,                   !- Maximum Regeneration Air Velocity for Humidity Ratio Equation m/s
    0.006911,                !- Minimum Regeneration Outlet Air Humidity Ratio for Humidity Ratio Equation kgWater/kgDryAir
    0.026707,                !- Maximum Regeneration Outlet Air Humidity Ratio for Humidity Ratio Equation kgWater/kgDryAir
    10.0,                    !- Minimum Regeneration Inlet Air Relative Humidity for Humidity Ratio Equation percent
    100.0,                   !- Maximum Regeneration Inlet Air Relative Humidity for Humidity Ratio Equation percent
    70.0,                    !- Minimum Process Inlet Air Relative Humidity for Humidity Ratio Equation percent
    100.0;                   !- Maximum Process Inlet Air Relative Humidity for Humidity Ratio Equation percent";

        private static IdfObject FindObject(IdfReader reader, string objectType, string objectName)
        {
            try
            {
                return reader[objectType].First(c => c[0] == objectName);
            }
            catch (Exception)
            {
                throw new Exception(string.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // ---------------------------
            // USER CONFIGURATION SECTION
            // ---------------------------
            string airLoopName = "Air Loop";         // Specify air loop name
            int insertAtBranchFieldIndex = 18;       // Branch field index where the new component is inserted.

            string desiccantHxObjectType = "HeatExchanger:Desiccant:BalancedFlow";
            string desiccantHxName = "Desiccant Unit";
            string mainBranchName = airLoopName + " AHU Main Branch";

            // Locate the Branch that represents the AHU main branch equipment list.
            IdfObject mainBranch = FindObject(idfReader, "Branch", mainBranchName);

            string processInletNodeName = mainBranch[insertAtBranchFieldIndex - 1].Value;
            string processOutletNodeName = "Desiccant Process Outlet Node";

            string existingHrHxName = airLoopName + " AHU Heat Recovery Device";
            IdfObject existingHeatRecoveryHx =
                FindObject(idfReader, "HeatExchanger:AirToAir:SensibleAndLatent", existingHrHxName);

            string regenerationInletNodeName = existingHeatRecoveryHx["Exhaust Air Outlet Node Name"].Value; // Return fan inlet node name
            string regenerationOutletNodeName = "Desiccant Return Outlet Node";

            string nextComponentType = mainBranch[insertAtBranchFieldIndex].Value;
            string nextComponentName = mainBranch[insertAtBranchFieldIndex + 1].Value;

            IdfObject nextComponent = FindObject(idfReader, nextComponentType, nextComponentName);
            nextComponent["Air Inlet Node Name"].Value = processOutletNodeName;
            mainBranch[insertAtBranchFieldIndex + 2].Value = processOutletNodeName;

            // Insert new Branch fields as a quad (Object Type, Object Name, Inlet Node, Outlet Node):
            string[] newBranchFields = new[]
            {
                desiccantHxObjectType,
                desiccantHxName,
                processInletNodeName,
                processOutletNodeName
            };
            mainBranch.InsertFields(insertAtBranchFieldIndex, newBranchFields);

            string desiccantHxIdfText = string.Format(
                desiccantHxAndPerfIdfTemplate,
                desiccantHxName,
                regenerationInletNodeName,
                regenerationOutletNodeName,
                processInletNodeName,
                processOutletNodeName);

            idfReader.Load(desiccantHxIdfText);
            idfReader.Save();
        }
    }
}