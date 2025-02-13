/*
Add HeatExchanger:Desiccant:BalancedFlow object into air handling unit.

The component is added at the specified index to the "DesiccantDehumidifier" air loop (the air loop name needs to match).

*/
using System.Runtime;
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
        string desiccantHxBoilerplate = @"     HeatExchanger:Desiccant:BalancedFlow,
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

        public IdfObject FindObject(IdfReader reader, string objectType, string objectName)
        {
            try
            {
                return reader[objectType].First(c => c[0] == objectName);
            }
            catch (Exception e)
            {
                throw new Exception(String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            string ahuName = "Air Loop";    // specify air loop name
            int dehumidfierBranchIndex = 18;    // place the component into the nth index in the AHU Main Branch

            string desiccantDehumidifierComponent = "HeatExchanger:Desiccant:BalancedFlow";
            string desiccantDehumidifierName = "Desiccant Unit";
            string ahuMainBranchName = ahuName + " AHU Main Branch";

            IdfObject ahuMainBranch = FindObject(idfReader, "Branch", ahuMainBranchName);


            string processInletNode = ahuMainBranch[dehumidfierBranchIndex - 1].Value;
            string processOutletNode = "Desiccant Process Outlet Node";

            string hxName = ahuName + " AHU Heat Recovery Device";
            IdfObject heatExchanger = FindObject(idfReader, "HeatExchanger:AirToAir:SensibleAndLatent", hxName);

            string regenerationInletNode = heatExchanger["Exhaust Air Outlet Node Name"].Value;    // return fan inlet node name
            string regenerationOutletNode = "Desiccant Return Outlet Node";

            string nextComponentType = ahuMainBranch[dehumidfierBranchIndex].Value;
            string nextComponentName = ahuMainBranch[dehumidfierBranchIndex + 1].Value;

            IdfObject nextComponent = FindObject(idfReader, nextComponentType, nextComponentName);
            nextComponent["Air Inlet Node Name"].Value = processOutletNode;
            ahuMainBranch[dehumidfierBranchIndex + 2].Value = processOutletNode;

            string[] newBranchFields = new string[] { desiccantDehumidifierComponent, desiccantDehumidifierName, processInletNode, processOutletNode };

            ahuMainBranch.InsertFields(dehumidfierBranchIndex, newBranchFields);

            string dehumidifier = String.Format(desiccantHxBoilerplate, desiccantDehumidifierName, regenerationInletNode, regenerationOutletNode, processInletNode, processOutletNode);

            idfReader.Load(dehumidifier);

            idfReader.Save();
        }
    }
}