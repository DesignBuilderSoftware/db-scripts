/*
Add Dehumidifier:Desiccant:System object into air handling unit.

The component is meant to be used in conjunction with a DX cooling coil.
The script looks up a DX coil coil and places the dehumidifer downstream of the coil.

The AddDesiccantDehumidifier method requires the following parameters:
    - air loop name

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
        string dehumidifierBoilerplate = @"  Dehumidifier:Desiccant:System,
    {0} Desiccant Dehumidifier,!- Name
    {0} Dehumidifier Schedule, !- Availability Schedule Name
    HeatExchanger:Desiccant:BalancedFlow,  !- Desiccant Heat Exchanger Object Type
    {0} Desiccant Heat Exchanger,  !- Desiccant Heat Exchanger Name
    {2},                     !- Sensor Node Name
    Fan:SystemModel,         !- Regeneration Air Fan Object Type
    {0} Desiccant Regen Fan, !- Regeneration Air Fan Name
    BlowThrough,             !- Regeneration Air Fan Placement
    ,                        !- Regeneration Air Heater Object Type
    ,                        !- Regeneration Air Heater Name
    46.111111,               !- Regeneration Inlet Air Setpoint Temperature C
    {3},                     !- Companion Cooling Coil Object Type
    {4},                     !- Companion Cooling Coil Name
    Yes,                     !- Companion Cooling Coil Upstream of Dehumidifier Process Inlet
    Yes,                     !- Companion Coil Regeneration Air Heating
    1.05,                    !- Exhaust Fan Maximum Flow Rate m3/s
    50,                      !- Exhaust Fan Maximum Power W
    {0} ExhaustFanPerfCurve; !- Exhaust Fan Power Curve Name

  HeatExchanger:Desiccant:BalancedFlow,
    {0} Desiccant Heat Exchanger,  !- Name
    {0} Dehumidifier Schedule,     !- Availability Schedule Name
    {0} Regen Fan Outlet Node,!- Regeneration Air Inlet Node Name
    {0} HX Regen Outlet Node,!- Regeneration Air Outlet Node Name
    {1},                     !- Process Air Inlet Node Name
    {2},                     !- Process Air Outlet Node Name
    HeatExchanger:Desiccant:BalancedFlow:PerformanceDataType1,  !- Heat Exchanger Performance Object Type
    HXDesPerf1;              !- Heat Exchanger Performance Name

  Fan:SystemModel,
    {0} Desiccant Regen Fan,     !- Name
    {0} Dehumidifier Schedule,   !- Availability Schedule Name
    {0} Outside Air Inlet Node 2,!- Air Inlet Node Name
    {0} Regen Fan Outlet Node,   !- Air Outlet Node Name
    Autosize,                    !- Design Maximum Air Flow Rate m3/s
    Continuous,              !- Speed Control Method
    0.0,                     !- Electric Power Minimum Flow Rate Fraction
    205.5,                   !- Design Pressure Rise Pa
    0.9,                     !- Motor Efficiency
    1.0,                     !- Motor In Air Stream Fraction
    AUTOSIZE,                !- Design Electric Power Consumption W
    TotalEfficiencyAndPressure,  !- Design Power Sizing Method
    ,                        !- Electric Power Per Unit Flow Rate W/(m3/s)
    ,                        !- Electric Power Per Unit Flow Rate Per Unit Pressure W/((m3/s)-Pa)
    0.7,                     !- Fan Total Efficiency
    {0} FanPerfCurve;        ! -  Electric Power Function of Flow Fraction Curve Name

 Curve:Quartic,
   {0} FanPerfCurve,         ! Curve Name
    0,                       ! CoefficientC1
    1,                       ! CoefficientC2
    0,                       ! CoefficientC3
    0,                       ! CoefficientC4
    0,                       ! CoefficientC5
    0,                       ! Minimum Value of x
    1,                       ! Maximum Value of x
    0,                       ! Minimum Curve Output
    1;                       ! Maximum Curve Output

  Curve:Cubic,
   {0} ExhaustFanPerfCurve,  !- Name
    0,                       !- Coefficient1 Constant
    1,                       !- Coefficient2 x
    0.0,                     !- Coefficient3 x**2
    0.0,                     !- Coefficient4 x**3
    0.0,                     !- Minimum Value of x
    1.0;                     !- Maximum Value of x  

OutdoorAir:Node,{0} Outside Air Inlet Node 2;

  Schedule:Compact, 
   {0} Dehumidifier Schedule,! Name
   Any Number,               ! Type
   Through: 12/31,           ! Type
   For: AllDays,             ! All days in year
   Until: 24:00,             ! All hours in day
   1;

  SetpointManager:MultiZone:Humidity:Maximum,
   {0} Humidity Setpoint Manager,                       ! - Component name
   {0},                                                 ! - HVAC air loop name
   .005,                                                ! - Minimum setpoint humidity ratio (kg/kg)
   .012,                                                ! - Maximum setpoint humidity ratio (kg/kg)
   {2};                                                 ! - Setpoint node list";

        string desiccantHeatExchangerPerf = @" HeatExchanger:Desiccant:BalancedFlow:PerformanceDataType1,
    HXDesPerf1,              !- Name
    1.05,                    !- Nominal Air Flow Rate {m3/s}
    3.25,                    !- Nominal Air Face Velocity {m/s}
    50.0,                    !- Nominal Electric Power {W}
    -2.53636E+00,            !- Temperature Equation Coefficient 1
    2.13247E+01,             !- Temperature Equation Coefficient 2
    9.23308E-01,             !- Temperature Equation Coefficient 3
    9.43276E+02,             !- Temperature Equation Coefficient 4
    -5.92367E+01,            !- Temperature Equation Coefficient 5
    -4.27465E-02,            !- Temperature Equation Coefficient 6
    1.12204E+02,             !- Temperature Equation Coefficient 7
    7.78252E-01,             !- Temperature Equation Coefficient 8
    0.007143,                !- Minimum Regeneration Inlet Air Humidity Ratio for Temperature Equation {kgWater/kgDryAir}
    0.024286,                !- Maximum Regeneration Inlet Air Humidity Ratio for Temperature Equation {kgWater/kgDryAir}
    46.111110,               !- Minimum Regeneration Inlet Air Temperature for Temperature Equation {C}
    46.111112,               !- Maximum Regeneration Inlet Air Temperature for Temperature Equation {C}
    0.005000,                !- Minimum Process Inlet Air Humidity Ratio for Temperature Equation {kgWater/kgDryAir}
    0.015714,                !- Maximum Process Inlet Air Humidity Ratio for Temperature Equation {kgWater/kgDryAir}
    4.583333,                !- Minimum Process Inlet Air Temperature for Temperature Equation {C}
    21.83333,                !- Maximum Process Inlet Air Temperature for Temperature Equation {C}
    2.286,                   !- Minimum Regeneration Air Velocity for Temperature Equation {m/s}
    4.826,                   !- Maximum Regeneration Air Velocity for Temperature Equation {m/s}
    35.0,                    !- Minimum Regeneration Outlet Air Temperature for Temperature Equation {C}
    50.0,                    !- Maximum Regeneration Outlet Air Temperature for Temperature Equation {C}
    5.0,                     !- Minimum Regeneration Inlet Air Relative Humidity for Temperature Equation {percent}
    45.0,                    !- Maximum Regeneration Inlet Air Relative Humidity for Temperature Equation {percent}
    80.0,                    !- Minimum Process Inlet Air Relative Humidity for Temperature Equation {percent}
    100.0,                   !- Maximum Process Inlet Air Relative Humidity for Temperature Equation {percent}
    -2.25547E+01,            !- Humidity Ratio Equation Coefficient 1
    9.76839E-01,             !- Humidity Ratio Equation Coefficient 2
    4.89176E-01,             !- Humidity Ratio Equation Coefficient 3
    -6.30019E-02,            !- Humidity Ratio Equation Coefficient 4
    1.20773E-02,             !- Humidity Ratio Equation Coefficient 5
    5.17134E-05,             !- Humidity Ratio Equation Coefficient 6
    4.94917E-02,             !- Humidity Ratio Equation Coefficient 7
    -2.59417E-04,            !- Humidity Ratio Equation Coefficient 8
    0.007143,                !- Minimum Regeneration Inlet Air Humidity Ratio for Humidity Ratio Equation {kgWater/kgDryAir}
    0.024286,                !- Maximum Regeneration Inlet Air Humidity Ratio for Humidity Ratio Equation {kgWater/kgDryAir}
    46.111110,               !- Minimum Regeneration Inlet Air Temperature for Humidity Ratio Equation {C}
    46.111112,               !- Maximum Regeneration Inlet Air Temperature for Humidity Ratio Equation {C}
    0.005000,                !- Minimum Process Inlet Air Humidity Ratio for Humidity Ratio Equation {kgWater/kgDryAir}
    0.015714,                !- Maximum Process Inlet Air Humidity Ratio for Humidity Ratio Equation {kgWater/kgDryAir}
    4.583333,                !- Minimum Process Inlet Air Temperature for Humidity Ratio Equation {C}
    21.83333,                !- Maximum Process Inlet Air Temperature for Humidity Ratio Equation {C}
    2.286,                   !- Minimum Regeneration Air Velocity for Humidity Ratio Equation {m/s}
    4.826,                   !- Maximum Regeneration Air Velocity for Humidity Ratio Equation {m/s}
    0.007914,                !- Minimum Regeneration Outlet Air Humidity Ratio for Humidity Ratio Equation {kgWater/kgDryAir}
    0.026279,                !- Maximum Regeneration Outlet Air Humidity Ratio for Humidity Ratio Equation {kgWater/kgDryAir}
    5.0,                     !- Minimum Regeneration Inlet Air Relative Humidity for Humidity Ratio Equation {percent}
    45.0,                    !- Maximum Regeneration Inlet Air Relative Humidity for Humidity Ratio Equation {percent}
    80.0,                    !- Minimum Process Inlet Air Relative Humidity for Humidity Ratio Equation {percent}
    100.0;                   !- Maximum Process Inlet Air Relative Humidity for Humidity Ratio Equation {percent}";

        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            AddDesiccantDehumidifier(idfReader, "Air Loop");
            idfReader.Load(desiccantHeatExchangerPerf);

            idfReader.Save();
        }

        public void AddDesiccantDehumidifier(IdfReader idfReader, string airLoopName)
        {
            string desiccantDehumidifierComponent = "Dehumidifier:Desiccant:System";
            string ahuMainBranchName = airLoopName + " AHU Main Branch";

            IdfObject ahuMainBranch = FindObject(idfReader, "Branch", ahuMainBranchName);

            int dxCoilSystemIndex = FindDxCoilSystemIndex(ahuMainBranch);
            if (dxCoilSystemIndex < 0)
            {
                throw new Exception("Cannot find DX coil in the AHU main branch.");
            }

            string dxCoilSystemType = ahuMainBranch[dxCoilSystemIndex].Value;
            string dxCoilSystemName = ahuMainBranch[dxCoilSystemIndex + 1].Value;

            IdfObject dxCoilSystem = FindObject(idfReader, dxCoilSystemType, dxCoilSystemName);
            string dxCoilType = dxCoilSystem["Cooling Coil Object Type"].Value;
            string dxCoilName = dxCoilSystem["Cooling Coil Name"].Value;

            IdfObject dxCoil = FindObject(idfReader, dxCoilType, dxCoilName);
            dxCoil["Condenser Air Inlet Node Name"].Value = airLoopName + " Outside Air Inlet Node 2";

            // update component nodes
            string processInletNode = "";
            string processOutletNode = "";

            int nextComponentIndex = dxCoilSystemIndex + 4;
            bool isLastComponent = nextComponentIndex == ahuMainBranch.Count;
            if (isLastComponent)
            {
                // update nodes
                processInletNode = airLoopName + " Desiccant Process Inlet Node";
                processOutletNode = ahuMainBranch[ahuMainBranch.Count - 1].Value;
                ahuMainBranch[dxCoilSystemIndex + 3].Value = processInletNode;

                dxCoilSystem["DX Cooling Coil System Outlet Node Name"].Value = processInletNode;
                dxCoil["Air Outlet Node Name"].Value = processInletNode;

                // append dehumidifier component into AHU main branch
                string[] newBranchFields = new string[] { desiccantDehumidifierComponent, airLoopName + " Desiccant Dehumidifier", processInletNode, processOutletNode };
                ahuMainBranch.AddFields(newBranchFields);
            }
            else
            {
                processInletNode = ahuMainBranch[dxCoilSystemIndex + 3].Value;
                processOutletNode = airLoopName + " Desiccant Process Outlet Node";

                // find the next component
                string nextComponentType = ahuMainBranch[nextComponentIndex].Value;
                string nextComponentName = ahuMainBranch[nextComponentIndex + 1].Value;

                // update nodes
                IdfObject nextComponent = FindObject(idfReader, nextComponentType, nextComponentName);
                nextComponent["Air Inlet Node Name"].Value = processOutletNode;
                ahuMainBranch[nextComponentIndex + 2].Value = processOutletNode;

                // insert dehumidifier component into AHU main branch
                string[] newBranchFields = new string[] { desiccantDehumidifierComponent, airLoopName + " Desiccant Dehumidifier", processInletNode, processOutletNode };
                ahuMainBranch.InsertFields(nextComponentIndex, newBranchFields);
            }

            // create object idf content
            string dehumidifier = String.Format(dehumidifierBoilerplate, airLoopName, processInletNode, processOutletNode, dxCoilType, dxCoilName);
            idfReader.Load(dehumidifier);
        }

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

        private int FindDxCoilSystemIndex(IdfObject mainBranch)
        {
            for (int i = 2; i < mainBranch.Count; i += 1)
            {
                if (mainBranch[i].Value.Equals("CoilSystem:Cooling:DX", StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }
            return -1;
        }
    }
}