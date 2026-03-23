/*
Add Desiccant Dehumidifier (type System) to an AirLoopHVAC main branch.

Purpose:
This DesignBuilder C# script inserts a Dehumidifier:Desiccant:System downstream of a DX cooling coil (Coilsystem:Cooling:DX) in an AHU main branch. 
It also loads the supporting heat exchanger, fan, curves, outdoor air nodes, schedule, and humidity setpoint manager objects.

Main Steps:
1) Find the target branch based on AHU name.
2) Locate the CoilSystem:Cooling:DX on that branch and resolve the referenced cooling coil object.
3) Re-wire nodes so the dehumidifier sits downstream of the DX coil system on the branch:
   - If DX coil system is the last component: append dehumidifier at end of branch.
   - If not last: insert dehumidifier and update the next component’s inlet node.
4) Load the required boilerplate IDF text blocks (dehumidifier, HX, fan, curves, schedule, OA nodes, setpoint manager)
   and load the HX performance data.

How to Use:

Configuration
Defined in the AddDesiccantDehumidifierToAirLoop():
- airLoopName: Must match the model's IDF naming.

Prerequisites / Placeholders
This script expects an Air Loop to be in place in the model defined in AddDesiccantDehumidifierToAirLoop().
The component is meant to be used in conjunction with a DX cooling coil.
The script looks up a DX coil coil and places the dehumidifer downstream of the coil.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. 
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime;
using System.Windows.Forms;

using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class AddDesiccantDehumidifier : ScriptBase, IScript
    {
        // IDF template that defines:
        // - Dehumidifier:Desiccant:System
        // - HeatExchanger:Desiccant:BalancedFlow
        // - Fan:SystemModel + curves
        // - OutdoorAir:Node placeholders
        // - Schedule + SetpointManager:MultiZone:Humidity:Maximum
        private readonly string dehumidifierIdfTemplate = @"  
Dehumidifier:Desiccant:System,
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
    {0} Outside Air Inlet Node 3,!- Air Inlet Node Name
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
OutdoorAir:Node,{0} Outside Air Inlet Node 3;

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

        // Performance data required by HeatExchanger:Desiccant:BalancedFlow
        private readonly string hxPerformanceDataIdf = @"
HeatExchanger:Desiccant:BalancedFlow:PerformanceDataType1,
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
            IdfReader idf = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // ----------------------------
            // USER CONFIGURATION SECTION
            // ----------------------------
            // Target air loop name must match the AirLoopHVAC name in the IDF.
            AddDesiccantDehumidifierToAirLoop(idf, "Air Loop");
            idf.Load(hxPerformanceDataIdf);

            idf.Save();
        }

        public void AddDesiccantDehumidifierToAirLoop(IdfReader idf, string airLoopName)
        {
            const string dehumidifierObjectType = "Dehumidifier:Desiccant:System";

            string ahuMainBranchName = airLoopName + " AHU Main Branch";
            IdfObject ahuMainBranch = FindObject(idf, "Branch", ahuMainBranchName);

            // Locate CoilSystem:Cooling:DX in the Branch object fields.
            int dxCoolingCoilSystemFieldIndex = FindDxCoilSystemIndex(ahuMainBranch);
            if (dxCoolingCoilSystemFieldIndex < 0)
            {
                throw new Exception("Cannot find DX coil (CoilSystem:Cooling:DX) in the AHU main branch.");
            }

            string dxCoilSystemType = ahuMainBranch[dxCoolingCoilSystemFieldIndex].Value;
            string dxCoilSystemName = ahuMainBranch[dxCoolingCoilSystemFieldIndex + 1].Value;

            IdfObject dxCoilSystem = FindObject(idf, dxCoilSystemType, dxCoilSystemName);
            string dxCoilType = dxCoilSystem["Cooling Coil Object Type"].Value;
            string dxCoilName = dxCoilSystem["Cooling Coil Name"].Value;

            IdfObject dxCoil = FindObject(idf, dxCoilType, dxCoilName);
            dxCoil["Condenser Air Inlet Node Name"].Value = airLoopName + " Outside Air Inlet Node 2";

            // Nodes between DX coil and dehumidifier:
            // - processInletNode = dehumidifier process inlet (downstream of DX coil outlet)
            // - processOutletNode = dehumidifier process outlet (upstream of next component or branch outlet)
            string processInletNode;
            string processOutletNode;

            int nextComponentFieldIndex = dxCoolingCoilSystemFieldIndex + 4;
            bool isDxCoilSystemLastOnBranch = nextComponentFieldIndex == ahuMainBranch.Count;

            if (isDxCoilSystemLastOnBranch)
            {
                // DX system is last: append dehumidifier and connect its outlet to the branch outlet node.
                processInletNode = airLoopName + " Desiccant Process Inlet Node";
                processOutletNode = ahuMainBranch[ahuMainBranch.Count - 1].Value;
                ahuMainBranch[dxCoolingCoilSystemFieldIndex + 3].Value = processInletNode;

                dxCoilSystem["DX Cooling Coil System Outlet Node Name"].Value = processInletNode;
                dxCoil["Air Outlet Node Name"].Value = processInletNode;

                // Append dehumidifier as a new component group in the branch.
                string[] newBranchFields = new string[]
                {
                    dehumidifierObjectType,
                    airLoopName + " Desiccant Dehumidifier",
                    processInletNode,
                    processOutletNode
                };
                ahuMainBranch.AddFields(newBranchFields);
            }
            else
            {
                // DX system is not last: insert dehumidifier and re-wire the next component inlet.
                processInletNode = ahuMainBranch[dxCoolingCoilSystemFieldIndex + 3].Value;
                processOutletNode = airLoopName + " Desiccant Process Outlet Node";

                // Identify the next component so we can update its inlet node to the dehumidifier outlet.
                string nextComponentType = ahuMainBranch[nextComponentFieldIndex].Value;
                string nextComponentName = ahuMainBranch[nextComponentFieldIndex + 1].Value;

                IdfObject nextComponent = FindObject(idf, nextComponentType, nextComponentName);
                nextComponent["Air Inlet Node Name"].Value = processOutletNode;

                // Update the branch’s “inlet node” field for the next component group.
                ahuMainBranch[nextComponentFieldIndex + 2].Value = processOutletNode;

                // Insert dehumidifier component group before the next component group.
                string[] newBranchFields = new string[]
                {
                    dehumidifierObjectType,
                    airLoopName + " Desiccant Dehumidifier",
                    processInletNode,
                    processOutletNode
                };
                ahuMainBranch.InsertFields(nextComponentFieldIndex, newBranchFields);
            }

            // Load the dehumidifier + associated objects into the IDF.
            string dehumidifierIdf = String.Format(
                dehumidifierIdfTemplate,
                airLoopName,
                processInletNode,
                processOutletNode,
                dxCoilType,
                dxCoilName);

            idf.Load(dehumidifierIdf);
        }

        public IdfObject FindObject(IdfReader reader, string objectType, string objectName)
        {
            try
            {
                return reader[objectType].First(c => c[0] == objectName);
            }
            catch
            {
                throw new Exception(String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        private int FindDxCoilSystemIndex(IdfObject mainBranch)
        {
            // Search across Branch fields for a CoilSystem:Cooling:DX entry.
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