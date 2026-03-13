/*
Add Desiccant Dehumidifier (type System) to an AirLoopHVAC outdoor air system (OA System).

This DesignBuilder C# script inserts a Dehumidifier:Desiccant:System into an air loop’s AirLoopHVAC:OutdoorAirSystem:EquipmentList, 
downstream of a CoilSystem:Cooling:DX component.
It also loads all supporting IDF objects required by the desiccant system plus a HX performance object.

Purpose
1) Find the CoilSystem:Cooling:DX entry in that OA equipment list.
2) Identify the companion DX cooling coil referenced by the DX coil system.
3) Rewire nodes so the new desiccant component sits downstream of the DX coil and upstream
  of the next OA component (special handling if the next component is OutdoorAir:Mixer).
4) Insert Dehumidifier:Desiccant:System into the OA equipment list and load the IDF text for the new objects.

How to Use

Configuration
Defined in the AddDesiccantDehumidifier():
- airLoopName: Must match the model's IDF naming.

Prerequisites (required placeholders)

This script expects an Air Loop to be in place in the model defined in AddDesiccantDehumidifier().
The component is meant to be used in conjunction with a DX cooling coil.
The script looks up a DX coil coil and places the dehumidifer downstream of the coil.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Linq;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class DesiccantDehumidifierOaSystemScript : ScriptBase, IScript
    {
        // IDF template for the full desiccant system "package" (dehumidifier + HX + regen fan + curves + schedule + OA nodes + setpoint managers).
        // Placeholders:
        //   {0} = airLoopName
        //   {1} = process inlet node name
        //   {2} = process outlet node name
        //   {3} = companion cooling coil object type
        //   {4} = companion cooling coil name
        private readonly string desiccantSystemIdfTemplate = @"
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
    34.0,                    !- Regeneration Inlet Air Setpoint Temperature C
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

SetpointManager:OutdoorAirPretreat,
  {0} Desiccant SetPoint Manager,  !- Name
   MaximumHumidityRatio,    !- Control Variable
  -99,                      !- Minimum Setpoint Temperature C
   99,                      !- Maximum Setpoint Temperature C
   0.00001,                 !- Minimum Setpoint Humidity Ratio kgWater/kgDryAir
   1.0,                     !- Maximum Setpoint Humidity Ratio kgWater/kgDryAir
   {0} Air Loop AHU Mixed Air Outlet,                 !- Reference Setpoint Node Name
   {0} Air Loop AHU Mixed Air Outlet,                 !- Mixed Air Stream Node Name
   {2},                                             !- Outdoor Air Stream Node Name
   {0} Air Loop AHU Extract Fan Air Outlet Node,      !- Return Air Stream Node Name
   {2};                                              !- Setpoint Node or NodeList Name

SetpointManager:MultiZone:Humidity:Maximum,
   {0} Maximum Mzone HUMRAT setpoint,!- Name
   {0} Air Loop,                     !- HVAC Air Loop Name
   0.001,                            !- Minimum Setpoint Humidity Ratio (kgWater/kgDryAir)
   0.057,                            !- Maximum Setpoint Humidity Ratio (kgWater/kgDryAir)
   {0} Air Loop AHU Mixed Air Outlet;!- Setpoint Node or NodeList Name";

        // HX performance object referenced by the desiccant heat exchanger above (name fixed as HXDesPerf1).
        private readonly string desiccantHxPerformanceIdf = @"HeatExchanger:Desiccant:BalancedFlow:PerformanceDataType1,
    HXDesPerf1,              !- Name
    1.16,                    !- Nominal Air Flow Rate {m3/s}
    3.24,                    !- Nominal Air Face Velocity {m/s}
    120,                     !- Nominal Electric Power {W}
    -2.53636E+00,            !- Temperature Equation Coefficient 1
    2.13247E+01,             !- Temperature Equation Coefficient 2
    9.23308E-01,             !- Temperature Equation Coefficient 3
    9.43276E+02,             !- Temperature Equation Coefficient 4
    -5.92367E+01,            !- Temperature Equation Coefficient 5
    -4.27465E-02,            !- Temperature Equation Coefficient 6
    1.12204E+02,             !- Temperature Equation Coefficient 7
    7.78252E-01,             !- Temperature Equation Coefficient 8
    0.001,                   !- Minimum Regeneration Inlet Air Humidity Ratio for Temperature Equation {kgWater/kgDryAir}
    0.0238,                  !- Maximum Regeneration Inlet Air Humidity Ratio for Temperature Equation {kgWater/kgDryAir}
    33.90,                   !- Minimum Regeneration Inlet Air Temperature for Temperature Equation {C}
    34.00,                   !- Maximum Regeneration Inlet Air Temperature for Temperature Equation {C}
    0.008,                   !- Minimum Process Inlet Air Humidity Ratio for Temperature Equation {kgWater/kgDryAir}
    0.00801,                 !- Maximum Process Inlet Air Humidity Ratio for Temperature Equation {kgWater/kgDryAir}
    10.62,                   !- Minimum Process Inlet Air Temperature for Temperature Equation {C}
    10.72,                   !- Maximum Process Inlet Air Temperature for Temperature Equation {C}
    3.14,                    !- Minimum Regeneration Air Velocity for Temperature Equation {m/s}
    3.24,                    !- Maximum Regeneration Air Velocity for Temperature Equation {m/s}
    44.9,                    !- Minimum Regeneration Outlet Air Temperature for Temperature Equation {C}
    45.0,                    !- Maximum Regeneration Outlet Air Temperature for Temperature Equation {C}
    69,                      !- Minimum Regeneration Inlet Air Relative Humidity for Temperature Equation {percent}
    70,                      !- Maximum Regeneration Inlet Air Relative Humidity for Temperature Equation {percent}
    99.8,                    !- Minimum Process Inlet Air Relative Humidity for Temperature Equation {percent}
    99.9,                    !- Maximum Process Inlet Air Relative Humidity for Temperature Equation {percent}
    -2.25547E+01,            !- Humidity Ratio Equation Coefficient 1
    9.76839E-01,             !- Humidity Ratio Equation Coefficient 2
    4.89176E-01,             !- Humidity Ratio Equation Coefficient 3
    -6.30019E-02,            !- Humidity Ratio Equation Coefficient 4
    1.20773E-02,             !- Humidity Ratio Equation Coefficient 5
    5.17134E-05,             !- Humidity Ratio Equation Coefficient 6
    4.94917E-02,             !- Humidity Ratio Equation Coefficient 7
    -2.59417E-04,            !- Humidity Ratio Equation Coefficient 8
    0.0238,                  !- Minimum Regeneration Inlet Air Humidity Ratio for Humidity Ratio Equation {kgWater/kgDryAir}
    0.0238001,               !- Maximum Regeneration Inlet Air Humidity Ratio for Humidity Ratio Equation {kgWater/kgDryAir}
    33.9,                    !- Minimum Regeneration Inlet Air Temperature for Humidity Ratio Equation {C}
    34.00,                   !- Maximum Regeneration Inlet Air Temperature for Humidity Ratio Equation {C}
    0.007,                   !- Minimum Process Inlet Air Humidity Ratio for Humidity Ratio Equation {kgWater/kgDryAir}
    0.008,                   !- Maximum Process Inlet Air Humidity Ratio for Humidity Ratio Equation {kgWater/kgDryAir}
    10.62,                   !- Minimum Process Inlet Air Temperature for Humidity Ratio Equation {C}
    10.720,                  !- Maximum Process Inlet Air Temperature for Humidity Ratio Equation {C}
    3.14,                    !- Minimum Regeneration Air Velocity for Humidity Ratio Equation {m/s}
    3.24,                    !- Maximum Regeneration Air Velocity for Humidity Ratio Equation {m/s}
    0.0228,                  !- Minimum Regeneration Outlet Air Humidity Ratio for Humidity Ratio Equation {kgWater/kgDryAir}
    0.02380,                 !- Maximum Regeneration Outlet Air Humidity Ratio for Humidity Ratio Equation {kgWater/kgDryAir}
    69,                      !- Minimum Regeneration Inlet Air Relative Humidity for Humidity Ratio Equation {percent}
    70.0,                    !- Maximum Regeneration Inlet Air Relative Humidity for Humidity Ratio Equation {percent}
    99.8,                    !- Minimum Process Inlet Air Relative Humidity for Humidity Ratio Equation {percent}
    99.9;                    !- Maximum Process Inlet Air Relative Humidity for Humidity Ratio Equation {percent}";

        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // ----------------------------
            // USER CONFIGURATION SECTION
            // ----------------------------
            // Configure target air loop(s) here:
            AddDesiccantDehumidifier(idfReader, "DOAS 1");
            idfReader.Load(desiccantHxPerformanceIdf);

            idfReader.Save();
        }

        public void AddDesiccantDehumidifier(IdfReader idfReader, string airLoopName)
        {
            const string desiccantSystemObjectType = "Dehumidifier:Desiccant:System";
            string oaEquipmentListName = airLoopName + " Air Loop AHU Outdoor air Equipment List";

            IdfObject oaEquipmentList = FindObject(idfReader, "AirLoopHVAC:OutdoorAirSystem:EquipmentList", oaEquipmentListName);

            // Locate the CoilSystem:Cooling:DX entry within the equipment list fields.
            int dxCoilSystemFieldIndex = FindDxCoilSystemIndex(oaEquipmentList);
            if (dxCoilSystemFieldIndex < 0)
            {
                throw new Exception("Cannot find DX coil system (CoilSystem:Cooling:DX) in the OA Equipment list.");
            }

            string dxCoilSystemObjectType = oaEquipmentList[dxCoilSystemFieldIndex].Value;
            string dxCoilSystemObjectName = oaEquipmentList[dxCoilSystemFieldIndex + 1].Value;

            IdfObject dxCoilSystem = FindObject(idfReader, dxCoilSystemObjectType, dxCoilSystemObjectName);
            string dxCoolingCoilObjectType = dxCoilSystem["Cooling Coil Object Type"].Value;
            string dxCoolingCoilObjectName = dxCoilSystem["Cooling Coil Name"].Value;

            IdfObject dxCoolingCoil = FindObject(idfReader, dxCoolingCoilObjectType, dxCoolingCoilObjectName);
            dxCoolingCoil["Condenser Air Inlet Node Name"].Value = airLoopName + " Outside Air Inlet Node 2";

            // The desiccant process inlet is the DX coil outlet. The process outlet is a new node inserted into the OA path.
            string processInletNodeName = dxCoolingCoil["Air Outlet Node Name"].Value;
            string processOutletNodeName = airLoopName + " Process Air Outlet Node";

            // Identify the component that currently follows the DX coil system in the OA equipment list.
            int downstreamComponentFieldIndex = dxCoilSystemFieldIndex + 2;
            string downstreamComponentObjectType = oaEquipmentList[downstreamComponentFieldIndex].Value;
            string downstreamComponentObjectName = oaEquipmentList[downstreamComponentFieldIndex + 1].Value;

            // Rewire the downstream component to take air from the new process outlet node (dehumidifier outlet).
            IdfObject downstreamComponent = FindObject(idfReader, downstreamComponentObjectType, downstreamComponentObjectName);

            if (downstreamComponent.IdfClass.Equals("OutdoorAir:Mixer", StringComparison.OrdinalIgnoreCase))
            {
                // OutdoorAir:Mixer uses "Outdoor Air Stream Node Name" for the OA stream inlet.
                downstreamComponent["Outdoor Air Stream Node Name"].Value = processOutletNodeName;
            }
            else
            {
                // Many OA components use a generic "Air Outlet Node Name" style field for inlet/outlet connectivity.
                downstreamComponent["Air Outlet Node Name"].Value = processOutletNodeName;
            }

            // Insert the dehumidifier into the OA equipment list immediately before the downstream component.
            string[] newEquipmentFields = new string[]
            {
                desiccantSystemObjectType,
                airLoopName + " Desiccant Dehumidifier"
            };
            oaEquipmentList.InsertFields(downstreamComponentFieldIndex, newEquipmentFields);

            // Load all supporting objects for the desiccant system (HX, fan, curves, schedule, OA nodes, SPMs).
            string desiccantSystemIdf = string.Format(
                desiccantSystemIdfTemplate,
                airLoopName,
                processInletNodeName,
                processOutletNodeName,
                dxCoolingCoilObjectType,
                dxCoolingCoilObjectName);

            idfReader.Load(desiccantSystemIdf);
        }

        public IdfObject FindObject(IdfReader reader, string objectType, string objectName)
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

        private int FindDxCoilSystemIndex(IdfObject oaEquipmentList)
        {
            // Searches for the CoilSystem:Cooling:DX object type within the OA equipment list fields.
            for (int i = 1; i < oaEquipmentList.Count; i += 1)
            {
                if (oaEquipmentList[i].Value.Equals("CoilSystem:Cooling:DX", StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }
            return -1;
        }
    }
}