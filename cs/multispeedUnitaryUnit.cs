/*
EnergyPlus Multi-Speed HVAC System Conversion Script
   This DesignBuilder C# script automatically converts single-speed AirLoopHVAC:UnitarySystem objects to multi-speed systems with discrete fan control.
   The script executes before EnergyPlus simulation runs and modifies the IDF file by replacing single-speed components with multi-speed equivalents.
   ​   
   Purpose
   The script transforms generic unitary air loops into multi-speed systems by:
   - Converting Fan:VariableVolume to Fan:SystemModel with discrete speed control
   - Replacing Coil:Cooling:DX:SingleSpeed with Coil:Cooling:DX:MultiSpeed
   - Adding UnitarySystemPerformance:Multispeed specifications
   ​
   Maintaining existing Coil:Heating:Fuel heating coils unchanged


   How to Use

   Configuration
   - Identify target systems: Systems are selected by matching the NameSuffix property (e.g., "10 ton" or "20 ton") at the end of the air loop name

   Configure speed specifications: Modify the TwoSpeedSpecification or FourSpeedSpecification structs in BeforeEnergySimulation() method:
   - NameSuffix: String identifier at end of system name
   - HalfSpeedFlowFraction: Air flow rate at 50% speed (2-speed)
   - HalfSpeedCapacityFraction: Cooling capacity at 50% speed (2-speed)
   - HalfSpeedFanPowerFraction: Fan power at 50% speed (2-speed)
   
   - Speed1/2/3FlowFraction: Air flow rates for speeds 1-3 (4-speed)
   - Speed1/2/3CapacityFraction: Cooling capacities for speeds 1-3 (4-speed)
   - Speed1/2/3FanPowerFraction: Fan power for speeds 1-3 (4-speed)
   
   Prerequisites
   Base model must contain AirLoopHVAC:UnitarySystem objects with:
   - Fan:VariableVolume supply fan
   - Coil:Cooling:DX:SingleSpeed cooling coil
   - Coil:Heating:Fuel heating coil
   
   Cooling coil must have numeric values for Rated Air Flow Rate and Gross Rated Total Cooling Capacity!
   
*/
using System.Runtime;
using System.Linq;
using System.Collections.Generic;
using System;
using System.Text;
using System.Windows.Forms;
using DB.Extensibility.Contracts;
using EpNet;


namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        // EpNET IdfReader instance
        public IdfReader Reader;

        public override void BeforeEnergySimulation()
        {
            Reader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);
            TwoSpeedSpecification twoSpeedSpec = new TwoSpeedSpecification
            {
                NameSuffix = "10 ton",
                HalfSpeedFlowFraction = 0.66,
                HalfSpeedCapacityFraction = 0.50,
                HalfSpeedFanPowerFraction = 0.51
            };
            UpdateToTwoSpeedUnitaryUnit(twoSpeedSpec);

            FourSpeedSpecification fourSpeedSpec = new FourSpeedSpecification
            {
                NameSuffix = "20 ton",
                Speed1FlowFraction = 0.25,
                Speed1CapacityFraction = 0.25,
                Speed1FanPowerFraction = 0.10,
                Speed2FlowFraction = 0.50,
                Speed2CapacityFraction = 0.50,
                Speed2FanPowerFraction = 0.25,
                Speed3FlowFraction = 0.75,
                Speed3CapacityFraction = 0.75,
                Speed3FanPowerFraction = 0.50
            };
            UpdateToFourSpeedUnitaryUnit(fourSpeedSpec);

            LoadCurves();
            Reader.Save();
        }

        public struct TwoSpeedSpecification
        {
            public string NameSuffix;
            public double HalfSpeedFlowFraction;
            public double HalfSpeedCapacityFraction;
            public double HalfSpeedFanPowerFraction;
        }

        public struct FourSpeedSpecification
        {
            public string NameSuffix;
            public double Speed1FlowFraction;
            public double Speed1CapacityFraction;
            public double Speed1FanPowerFraction;
            public double Speed2FlowFraction;
            public double Speed2CapacityFraction;
            public double Speed2FanPowerFraction;
            public double Speed3FlowFraction;
            public double Speed3CapacityFraction;
            public double Speed3FanPowerFraction;
        }

        public void LoadCurves()
        {
            string curves = @"
  Curve:Quadratic,
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_CapFF,  !- Name
    0.77136,                 !- Coefficient1 Constant
    0.34053,                 !- Coefficient2 x
    -0.11088,                !- Coefficient3 x**2
    0.75918,                 !- Minimum Value of x
    1.13877;                 !- Maximum Value of x

  Curve:Quadratic,
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_EIRFFF,  !- Name
    1.20550,                 !- Coefficient1 Constant
    -0.32953,                !- Coefficient2 x
    0.12308,                 !- Coefficient3 x**2
    0.75918,                 !- Minimum Value of x
    1.13877;                 !- Maximum Value of x

  Curve:Quadratic,
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_PLR,  !- Name
    0.77100,                 !- Coefficient1 Constant
    0.22900,                 !- Coefficient2 x
    0.0,                     !- Coefficient3 x**2
    0.0,                     !- Minimum Value of x
    1.0;                     !- Maximum Value of x

  Curve:Biquadratic,
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_CapFT,  !- Name
    0.42415,                 !- Coefficient1 Constant
    0.04426,                 !- Coefficient2 x
    -0.00042,                !- Coefficient3 x**2
    0.00333,                 !- Coefficient4 y
    -0.00008,                !- Coefficient5 y**2
    -0.00021,                !- Coefficient6 x*y
    17.00000,                !- Minimum Value of x
    22.00000,                !- Maximum Value of x
    29.00000,                !- Minimum Value of y
    46.00000;                !- Maximum Value of y

  Curve:Biquadratic,
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_EIRFT,  !- Name
    1.23649,                 !- Coefficient1 Constant
    -0.02431,                !- Coefficient2 x
    0.00057,                 !- Coefficient3 x**2
    -0.01434,                !- Coefficient4 y
    0.00063,                 !- Coefficient5 y**2
    -0.00038,                !- Coefficient6 x*y
    17.00000,                !- Minimum Value of x
    22.00000,                !- Maximum Value of x
    29.00000,                !- Minimum Value of y
    46.00000;                !- Maximum Value of y
";
            Reader.Load(curves);
        }

        public void UpdateToTwoSpeedUnitaryUnit(TwoSpeedSpecification spec)
        {
            foreach (IdfObject unitarySystem in Reader["AirLoopHVAC:UnitarySystem"])
            {
                string systemName = unitarySystem["Name"].Value;
                string airLoopName = systemName.Replace(" Unitary System", "");
                if (airLoopName.EndsWith(spec.NameSuffix))
                {
                    string multispeedSpecName = airLoopName + " Multispeed Specification"; ;
                    string multispeedSpec = Apply2SpeedSpec(multispeedSpecName, unitarySystem, spec);

                    // Add multispeed specification to the unitary system
                    unitarySystem["Design Specification Multispeed Object Type"].Value = "UnitarySystemPerformance:Multispeed";
                    unitarySystem["Design Specification Multispeed Object Name"].Value = multispeedSpecName;

                    Reader.Load(multispeedSpec);
                }
            }
        }

        public void UpdateToFourSpeedUnitaryUnit(FourSpeedSpecification spec)
        {
            foreach (IdfObject unitarySystem in Reader["AirLoopHVAC:UnitarySystem"])
            {
                string systemName = unitarySystem["Name"].Value;
                string airLoopName = systemName.Replace(" Unitary System", "");
                if (airLoopName.EndsWith(spec.NameSuffix))
                {
                    string multispeedSpecName = airLoopName + " Multispeed Specification"; ;
                    string multispeedSpec = Apply4SpeedSpec(multispeedSpecName, unitarySystem, spec);

                    // Add multispeed specification to the unitary system
                    unitarySystem["Design Specification Multispeed Object Type"].Value = "UnitarySystemPerformance:Multispeed";
                    unitarySystem["Design Specification Multispeed Object Name"].Value = multispeedSpecName;

                    Reader.Load(multispeedSpec);
                }
            }
        }
        public string Apply2SpeedSpec(string multispeedSpecName, IdfObject unitarySystem, TwoSpeedSpecification spec)
        {
            StringBuilder multispeedIdfContent = new StringBuilder();

            string systemName = unitarySystem["Name"].Value;

            // Add multispeed specification
            string specification = Get2SMultispeedSpecification(multispeedSpecName, systemName, spec);
            multispeedIdfContent.Append(specification);

            // Update supply fan
            IdfObject fan = FindObject(unitarySystem["Supply Fan Object Type"].Value, unitarySystem["Supply Fan Name"].Value);
            string fanName = fan["Name"].Value;
            string newFan = Get2SFan(fan, spec);
            ReplaceObjectTypeInList(
                "AirLoopHVAC:UnitarySystem",
                "Fan:VariableVolume",
                fanName,
                "Fan:SystemModel",
                fanName);
            Reader.Load(newFan);
            Reader.Remove(fan);

            // Update cooling coil
            IdfObject coolingCoil = FindObject(unitarySystem["Cooling Coil Object Type"].Value, unitarySystem["Cooling Coil Name"].Value);
            if (coolingCoil.IdfClass == "Coil:Cooling:DX:SingleSpeed")
            {
                string multispeedCoolingCoil = Get2SCoolingCoil(coolingCoil, spec);
                string multispeedCoolingCoilName = coolingCoil["Name"].Value;
                ReplaceObjectTypeInList(
                    "AirLoopHVAC:UnitarySystem",
                    "Coil:Cooling:DX:SingleSpeed",
                    multispeedCoolingCoilName,
                    "Coil:Cooling:DX:MultiSpeed",
                    multispeedCoolingCoilName);
                Reader.Load(multispeedCoolingCoil);
                Reader.Remove(coolingCoil);
            }
            else
            {
                throw new Exception(
                    "Cannot convert cooling coil to 2-speed, the required placeholder object is Coil:Cooling:DX:SingleSpeed, system: " + systemName);
            }

            // Update heating coil
            IdfObject heatingCoil = FindObject(unitarySystem["Heating Coil Object Type"].Value, unitarySystem["Heating Coil Name"].Value);
            if (heatingCoil.IdfClass == "Coil:Heating:Fuel")
            {
                // Leave the fuel heating coil as is
            }
            else
            {
                throw new Exception(
                    "Cannot update system: " + systemName + ", the only valid option for heating coil is currently Coil:Heating:Fuel.");
            }

            return multispeedIdfContent.ToString();
        }
        public string Apply4SpeedSpec(string multispeedSpecName, IdfObject unitarySystem, FourSpeedSpecification spec)
        {
            StringBuilder multispeedIdfContent = new StringBuilder();

            string systemName = unitarySystem["Name"].Value;

            // Add multispeed specification
            string specification = Get4SMultispeedSpecification(multispeedSpecName, systemName, spec);
            multispeedIdfContent.Append(specification);

            // Update supply fan
            IdfObject fan = FindObject(unitarySystem["Supply Fan Object Type"].Value, unitarySystem["Supply Fan Name"].Value);
            string fanName = fan["Name"].Value;
            string newFan = Get4SFan(fan, spec);
            ReplaceObjectTypeInList(
                "AirLoopHVAC:UnitarySystem",
                "Fan:VariableVolume",
                fanName,
                "Fan:SystemModel",
                fanName);
            Reader.Load(newFan);
            Reader.Remove(fan);

            // Update cooling coil
            IdfObject coolingCoil = FindObject(unitarySystem["Cooling Coil Object Type"].Value, unitarySystem["Cooling Coil Name"].Value);
            if (coolingCoil.IdfClass == "Coil:Cooling:DX:SingleSpeed")
            {
                string multispeedCoolingCoil = Get4SCoolingCoil(coolingCoil, spec);
                string multispeedCoolingCoilName = coolingCoil["Name"].Value;
                ReplaceObjectTypeInList(
                    "AirLoopHVAC:UnitarySystem",
                    "Coil:Cooling:DX:SingleSpeed",
                    multispeedCoolingCoilName,
                    "Coil:Cooling:DX:MultiSpeed",
                    multispeedCoolingCoilName);
                Reader.Load(multispeedCoolingCoil);
                Reader.Remove(coolingCoil);
            }
            else
            {
                throw new Exception(
                    "Cannot convert cooling coil to 4-speed, the required placeholder object is Coil:Cooling:DX:SingleSpeed, system: " + systemName);
            }

            // Update heating coil
            IdfObject heatingCoil = FindObject(unitarySystem["Heating Coil Object Type"].Value, unitarySystem["Heating Coil Name"].Value);
            if (heatingCoil.IdfClass == "Coil:Heating:Fuel")
            {
                // Leave the fuel heating coil as is
            }
            else
            {
                throw new Exception(
                    "Cannot update system: " + systemName + ", the only valid option for heating coil is currently Coil:Heating:Fuel.");
            }

            return multispeedIdfContent.ToString();
        }

        public string Get2SMultispeedSpecification(string multispeedSpecName, string unitName, TwoSpeedSpecification spec)
        {
            string template = @"  UnitarySystemPerformance:Multispeed,
    {0},                     !- Name
    2,                       !- Number of Speeds for Heating
    2,                       !- Number of Speeds for Cooling
    Yes,                     !- Single Mode Operation
    0.0,                     !- No Load Supply Air Flow Rate Ratio
    {1},                     !- Heating Speed 1 Supply Air Flow Ratio
    {1},                     !- Cooling Speed 1 Supply Air Flow Ratio
    1,                       !- Heating Speed 2 Supply Air Flow Ratio
    1;                       !- Cooling Speed 2 Supply Air Flow Ratio";
            return string.Format(template, multispeedSpecName, spec.HalfSpeedFlowFraction);
        }

        public string Get4SMultispeedSpecification(string multispeedSpecName, string unitName, FourSpeedSpecification spec)
        {
            string template = @"  UnitarySystemPerformance:Multispeed,
    {0},                     !- Name
    4,                       !- Number of Speeds for Heating
    4,                       !- Number of Speeds for Cooling
    Yes,                     !- Single Mode Operation
    0.0,                     !- No Load Supply Air Flow Rate Ratio
    {1},                     !- Heating Speed 1 Supply Air Flow Ratio
    {1},                     !- Cooling Speed 1 Supply Air Flow Ratio
    {2},                     !- Heating Speed 2 Supply Air Flow Ratio
    {2},                     !- Cooling Speed 2 Supply Air Flow Ratio
    {3},                     !- Heating Speed 3 Supply Air Flow Ratio
    {3},                     !- Cooling Speed 3 Supply Air Flow Ratio
    1,                       !- Heating Speed 4 Supply Air Flow Ratio
    1;                       !- Cooling Speed 4 Supply Air Flow Ratio";
            return string.Format(template, multispeedSpecName, spec.Speed1FlowFraction, spec.Speed2FlowFraction, spec.Speed3FlowFraction);
        }

        public string Get2SFan(IdfObject fan, TwoSpeedSpecification spec)
        {
            if (fan.IdfClass != "Fan:VariableVolume")
            {
                throw new Exception(
                    "Cannot get fan performance, the required placeholder object is Fan:VariableVolume., fan: " + fan["Name"].Value);
            }
            string newFanTemplate = @"  Fan:SystemModel,
  {0},           !- Name
  {1},           !- Availability Schedule Name
  {2},           !- Air Inlet Node Name
  {3},           !- Air Outlet Node Name
  {4},           !- Design Maximum Air Flow Rate
  Discrete ,     !- Speed Control Method
  0.0,           !- Electric Power Minimum Flow Rate Fraction
  {5},           !- Design Pressure Rise
  {6},           !- Motor Efficiency
  {7},           !- Motor In Air Stream Fraction
  AUTOSIZE,      !- Design Electric Power Consumption
  TotalEfficiencyAndPressure, !- Design Power Sizing Method
  ,              !- Electric Power Per Unit Flow Rate
  ,              !- Electric Power Per Unit Flow Rate Per Unit Pressure
  {8},           !- Fan Total Efficiency
  ,              !- Electric Power Function of Flow Fraction Curve Name
  ,              !- Night Ventilation Mode Pressure Rise
  ,              !- Night Ventilation Mode Flow Fraction
  ,              !- Motor Loss Zone Name
  ,              !- Motor Loss Radiative Fraction
  {9},           !- End-Use Subcategory
  2,             !- Number of Speeds
  {10},          !- Speed 1 Flow Fraction
  {11},          !- Speed 1 Electric Power Fraction
  1.0,           !- Speed 2 Flow Fraction
  1.0;           !- Speed 2 Electric Power Fraction";

            return String.Format(
                newFanTemplate,
                fan[0].Value,
                fan[1].Value,
                fan[15].Value,
                fan[16].Value,
                fan[4].Value,
                fan[3].Value,
                fan[8].Value,
                fan[9].Value,
                fan[2].Value,
                fan[17].Value,
                spec.HalfSpeedFlowFraction,
                spec.HalfSpeedFanPowerFraction);
        }

        public string Get4SFan(IdfObject fan, FourSpeedSpecification spec)
        {
            if (fan.IdfClass != "Fan:VariableVolume")
            {
                throw new Exception(
                    "Cannot get fan performance, the required placeholder object is Fan:VariableVolume., fan: " + fan["Name"].Value);
            }
            string newFanTemplate = @"  Fan:SystemModel,
  {0},           !- Name
  {1},           !- Availability Schedule Name
  {2},           !- Air Inlet Node Name
  {3},           !- Air Outlet Node Name
  {4},           !- Design Maximum Air Flow Rate
  Discrete ,     !- Speed Control Method
  0.0,           !- Electric Power Minimum Flow Rate Fraction
  {5},           !- Design Pressure Rise
  {6},           !- Motor Efficiency
  {7},           !- Motor In Air Stream Fraction
  AUTOSIZE,      !- Design Electric Power Consumption
  TotalEfficiencyAndPressure, !- Design Power Sizing Method
  ,              !- Electric Power Per Unit Flow Rate
  ,              !- Electric Power Per Unit Flow Rate Per Unit Pressure
  {8},           !- Fan Total Efficiency
  ,              !- Electric Power Function of Flow Fraction Curve Name
  ,              !- Night Ventilation Mode Pressure Rise
  ,              !- Night Ventilation Mode Flow Fraction
  ,              !- Motor Loss Zone Name
  ,              !- Motor Loss Radiative Fraction
  {9},           !- End-Use Subcategory
  4,             !- Number of Speeds
  {10},          !- Speed 1 Flow Fraction
  {11},          !- Speed 1 Electric Power Fraction
  {12},          !- Speed 2 Flow Fraction
  {13},          !- Speed 2 Electric Power Fraction
  {14},          !- Speed 3 Flow Fraction
  {15},          !- Speed 3 Electric Power Fraction
  1.0,           !- Speed 4 Flow Fraction
  1.0;           !- Speed 4 Electric Power Fraction";

            string newFan = String.Format(
                newFanTemplate,
                fan[0].Value,
                fan[1].Value,
                fan[15].Value,
                fan[16].Value,
                fan[4].Value,
                fan[3].Value,
                fan[8].Value,
                fan[9].Value,
                fan[2].Value,
                fan[17].Value,
                spec.Speed1FlowFraction,
                spec.Speed1FanPowerFraction,
                spec.Speed2FlowFraction,
                spec.Speed2FanPowerFraction,
                spec.Speed3FlowFraction,
                spec.Speed3FanPowerFraction);
            return newFan;
        }

        public string Get2SCoolingCoil(IdfObject coil, TwoSpeedSpecification spec)
        {
            string name = coil[0].Value;
            double ratedAirFlowRate = ConvertWithError(coil["Rated Air Flow Rate"].Value, name, "Rated Air Flow Rate");
            double ratedCapacity = ConvertWithError(coil["Gross Rated Total Cooling Capacity"].Value, name, "Rated total cooling capacity");

            double halfSpeedAirFlowRate = ratedAirFlowRate * spec.HalfSpeedFlowFraction;
            double halfSpeedCapacity = ratedCapacity * spec.HalfSpeedCapacityFraction;

            string cop = coil["Gross Rated Cooling COP"].Value;
            string shr = coil["Gross Rated Sensible Heat Ratio"].Value;

            string template = @"
  Coil:Cooling:DX:MultiSpeed,
    {0},                     !- Name
    {1},                     !- Availability Schedule Name
    {2},                     !- Air Inlet Node Name
    {3},                     !- Air Outlet Node Name
    ,                        !- Condenser Air Inlet Node Name
    AirCooled,               !- Condenser Type
    ,                        !- Minimum Outdoor Dry-Bulb Temperature for Compressor Operation [C]
    ,                        !- Supply Water Storage Tank Name
    ,                        !- Condensate Collection Water Storage Tank Name
    No,                      !- Apply Part Load Fraction to Speeds Greater than 1
    No,                      !- Apply Latent Degradation to Speeds Greater than 1
    {4},                     !- Crankcase Heater Capacity [W]
    {5},                     !- Maximum Outdoor Dry-Bulb Temperature for Crankcase Heater Operation [C]
    ,                        !- Basin Heater Capacity [W/K]
    2,                       !- Basin Heater Setpoint Temperature [C]
    ,                        !- Basin Heater Operating Schedule Name
    Electricity,             !- Fuel Type
    2,                       !- Number of Speeds
    {6},                     !- Speed 1 Gross Rated Total Cooling Capacity [W]
    {7},                     !- Speed 1 Gross Rated Sensible Heat Ratio
    {8},                     !- Speed 1 Gross Rated Cooling COP [W/W]
    {9},                     !- Speed 1 Rated Air Flow Rate [m3/s]
    {12},                    !- Speed 1 2017 Rated Evaporator Fan Power Per Volume Flow Rate [W/(m3/s)]
    {13},                    !- Speed 1 2023 Rated Evaporator Fan Power Per Volume Flow [W/(m3/s)]
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_CapFT,  !- Speed 1 Total Cooling Capacity Function of Temperature Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_CapFF,  !- Speed 1 Total Cooling Capacity Function of Flow Fraction Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_EIRFT,  !- Speed 1 Energy Input Ratio Function of Temperature Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_EIRFFF,  !- Speed 1 Energy Input Ratio Function of Flow Fraction Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_PLR,  !- Speed 1 Part Load Fraction Correlation Curve Name
    ,                        !- Speed 1 Nominal Time for Condensate Removal to Begin [s]
    ,                        !- Speed 1 Ratio of Initial Moisture Evaporation Rate and Steady State Latent Capacity [dimensionless]
    ,                        !- Speed 1 Maximum Cycling Rate [cycles/hr]
    ,                        !- Speed 1 Latent Capacity Time Constant [s]
    0.2,                     !- Speed 1 Rated Waste Heat Fraction of Power Input [dimensionless]
    ,                        !- Speed 1 Waste Heat Function of Temperature Curve Name
    0.9,                     !- Speed 1 Evaporative Condenser Effectiveness [dimensionless]
    ,                        !- Speed 1 Evaporative Condenser Air Flow Rate [m3/s]
    ,                        !- Speed 1 Rated Evaporative Condenser Pump Power Consumption [W]
    {10},                    !- Speed 2 Gross Rated Total Cooling Capacity [W]
    {7},                     !- Speed 2 Gross Rated Sensible Heat Ratio
    {8},                     !- Speed 2 Gross Rated Cooling COP [W/W]
    {11},                    !- Speed 2 Rated Air Flow Rate [m3/s]
    {12},                    !- Speed 2 2017 Rated Evaporator Fan Power Per Volume Flow Rate [W/(m3/s)]
    {13},                    !- Speed 2 2023 Rated Evaporator Fan Power Per Volume Flow [W/(m3/s)]
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_CapFT,  !- Speed 1 Total Cooling Capacity Function of Temperature Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_CapFF,  !- Speed 1 Total Cooling Capacity Function of Flow Fraction Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_EIRFT,  !- Speed 1 Energy Input Ratio Function of Temperature Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_EIRFFF,  !- Speed 1 Energy Input Ratio Function of Flow Fraction Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_PLR,  !- Speed 1 Part Load Fraction Correlation Curve Name
    ,                        !- Speed 2 Nominal Time for Condensate Removal to Begin [s]
    ,                        !- Speed 2 Ratio of Initial Moisture Evaporation Rate and steady state Latent Capacity [dimensionless]
    ,                        !- Speed 2 Maximum Cycling Rate [cycles/hr]
    ,                        !- Speed 2 Latent Capacity Time Constant [s]
    0.2,                     !- Speed 2 Rated Waste Heat Fraction of Power Input [dimensionless]
    ,                        !- Speed 2 Waste Heat Function of Temperature Curve Name
    0.9,                     !- Speed 2 Evaporative Condenser Effectiveness [dimensionless]
    ,                        !- Speed 2 Evaporative Condenser Air Flow Rate [m3/s]
    ;                        !- Speed 2 Rated Evaporative Condenser Pump Power Consumption [W]";
            return string.Format(
                template,
                coil["Name"].Value,
                coil["Availability Schedule Name"].Value,
                coil["Air Inlet Node Name"].Value,
                coil["Air Outlet Node Name"].Value,
                coil["Crankcase Heater Capacity"].Value,
                coil["Maximum Outdoor Dry-Bulb Temperature for Crankcase Heater Operation"].Value,
                halfSpeedCapacity,
                shr,
                cop,
                halfSpeedAirFlowRate,
                ratedCapacity,
                ratedAirFlowRate,
                coil["2017 Rated Evaporator Fan Power Per Volume Flow Rate"].Value,
                coil["2023 Rated Evaporator Fan Power Per Volume Flow Rate"].Value
                );
        }

        public string Get4SCoolingCoil(IdfObject coil, FourSpeedSpecification spec)
        {
            string name = coil[0].Value;
            double ratedAirFlowRate = ConvertWithError(coil["Rated Air Flow Rate"].Value, name, "Rated Air Flow Rate");
            double ratedCapacity = ConvertWithError(coil["Gross Rated Total Cooling Capacity"].Value, name, "Rated total cooling capacity");

            double speed1AirFlowRate = ratedAirFlowRate * spec.Speed1FlowFraction;
            double speed1Capacity = ratedCapacity * spec.Speed1CapacityFraction;
            double speed2AirFlowRate = ratedAirFlowRate * spec.Speed2FlowFraction;
            double speed2Capacity = ratedCapacity * spec.Speed2CapacityFraction;
            double speed3AirFlowRate = ratedAirFlowRate * spec.Speed3FlowFraction;
            double speed3Capacity = ratedCapacity * spec.Speed3CapacityFraction;

            string cop = coil["Gross Rated Cooling COP"].Value;
            string shr = coil["Gross Rated Sensible Heat Ratio"].Value;

            string template = @"
  Coil:Cooling:DX:MultiSpeed,
    {0},                     !- Name
    {1},                     !- Availability Schedule Name
    {2},                     !- Air Inlet Node Name
    {3},                     !- Air Outlet Node Name
    ,                        !- Condenser Air Inlet Node Name
    AirCooled,               !- Condenser Type
    ,                        !- Minimum Outdoor Dry-Bulb Temperature for Compressor Operation [C]
    ,                        !- Supply Water Storage Tank Name
    ,                        !- Condensate Collection Water Storage Tank Name
    No,                      !- Apply Part Load Fraction to Speeds Greater than 1
    No,                      !- Apply Latent Degradation to Speeds Greater than 1
    {4},                     !- Crankcase Heater Capacity [W]
    {5},                     !- Maximum Outdoor Dry-Bulb Temperature for Crankcase Heater Operation [C]
    ,                        !- Basin Heater Capacity [W/K]
    2,                       !- Basin Heater Setpoint Temperature [C]
    ,                        !- Basin Heater Operating Schedule Name
    Electricity,             !- Fuel Type
    4,                       !- Number of Speeds
    {10},                    !- Speed 1 Gross Rated Total Cooling Capacity [W]
    {8},                     !- Speed 1 Gross Rated Sensible Heat Ratio
    {9},                     !- Speed 1 Gross Rated Cooling COP [W/W]
    {11},                    !- Speed 1 Rated Air Flow Rate [m3/s]
    {6},                     !- Speed 1 2017 Rated Evaporator Fan Power Per Volume Flow Rate [W/(m3/s)]
    {7},                     !- Speed 1 2023 Rated Evaporator Fan Power Per Volume Flow [W/(m3/s)]
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_CapFT,  !- Speed 1 Total Cooling Capacity Function of Temperature Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_CapFF,  !- Speed 1 Total Cooling Capacity Function of Flow Fraction Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_EIRFT,  !- Speed 1 Energy Input Ratio Function of Temperature Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_EIRFFF,  !- Speed 1 Energy Input Ratio Function of Flow Fraction Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_PLR,  !- Speed 1 Part Load Fraction Correlation Curve Name
    ,                        !- Speed 1 Nominal Time for Condensate Removal to Begin [s]
    ,                        !- Speed 1 Ratio of Initial Moisture Evaporation Rate and Steady State Latent Capacity [dimensionless]
    ,                        !- Speed 1 Maximum Cycling Rate [cycles/hr]
    ,                        !- Speed 1 Latent Capacity Time Constant [s]
    0.2,                     !- Speed 1 Rated Waste Heat Fraction of Power Input [dimensionless]
    ,                        !- Speed 1 Waste Heat Function of Temperature Curve Name
    0.9,                     !- Speed 1 Evaporative Condenser Effectiveness [dimensionless]
    ,                        !- Speed 1 Evaporative Condenser Air Flow Rate [m3/s]
    ,                        !- Speed 1 Rated Evaporative Condenser Pump Power Consumption [W]
    {12},                    !- Speed 2 Gross Rated Total Cooling Capacity [W]
    {8},                     !- Speed 2 Gross Rated Sensible Heat Ratio
    {9},                     !- Speed 2 Gross Rated Cooling COP [W/W]
    {13},                    !- Speed 2 Rated Air Flow Rate [m3/s]
    {6},                     !- Speed 2 2017 Rated Evaporator Fan Power Per Volume Flow Rate [W/(m3/s)]
    {7},                     !- Speed 2 2023 Rated Evaporator Fan Power Per Volume Flow [W/(m3/s)]
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_CapFT,  !- Speed 2 Total Cooling Capacity Function of Temperature Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_CapFF,  !- Speed 2 Total Cooling Capacity Function of Flow Fraction Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_EIRFT,  !- Speed 2 Energy Input Ratio Function of Temperature Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_EIRFFF,  !- Speed 2 Energy Input Ratio Function of Flow Fraction Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_PLR,  !- Speed 2 Part Load Fraction Correlation Curve Name
    ,                        !- Speed 2 Nominal Time for Condensate Removal to Begin [s]
    ,                        !- Speed 2 Ratio of Initial Moisture Evaporation Rate and steady state Latent Capacity [dimensionless]
    ,                        !- Speed 2 Maximum Cycling Rate [cycles/hr]
    ,                        !- Speed 2 Latent Capacity Time Constant [s]
    0.2,                     !- Speed 2 Rated Waste Heat Fraction of Power Input [dimensionless]
    ,                        !- Speed 2 Waste Heat Function of Temperature Curve Name
    0.9,                     !- Speed 2 Evaporative Condenser Effectiveness [dimensionless]
    ,                        !- Speed 2 Evaporative Condenser Air Flow Rate [m3/s]
    ,                        !- Speed 2 Rated Evaporative Condenser Pump Power Consumption [W]
    {14},                    !- Speed 3 Gross Rated Total Cooling Capacity [W]
    {8},                     !- Speed 3 Gross Rated Sensible Heat Ratio
    {9},                     !- Speed 3 Gross Rated Cooling COP [W/W]
    {15},                    !- Speed 3 Rated Air Flow Rate [m3/s]
    {6},                     !- Speed 3 2017 Rated Evaporator Fan Power Per Volume Flow Rate [W/(m3/s)]
    {7},                     !- Speed 3 2023 Rated Evaporator Fan Power Per Volume Flow [W/(m3/s)]
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_CapFT,  !- Speed 3 Total Cooling Capacity Function of Temperature Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_CapFF,  !- Speed 3 Total Cooling Capacity Function of Flow Fraction Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_EIRFT,  !- Speed 3 Energy Input Ratio Function of Temperature Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_EIRFFF,  !- Speed 3 Energy Input Ratio Function of Flow Fraction Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_PLR,  !- Speed 3 Part Load Fraction Correlation Curve Name
    ,                        !- Speed 3 Nominal Time for Condensate Removal to Begin [s]
    ,                        !- Speed 3 Ratio of Initial Moisture Evaporation Rate and steady state Latent Capacity [dimensionless]
    ,                        !- Speed 3 Maximum Cycling Rate [cycles/hr]
    ,                        !- Speed 3 Latent Capacity Time Constant [s]
    0.2,                     !- Speed 3 Rated Waste Heat Fraction of Power Input [dimensionless]
    ,                        !- Speed 3 Waste Heat Function of Temperature Curve Name
    0.9,                     !- Speed 3 Evaporative Condenser Effectiveness [dimensionless]
    ,                        !- Speed 3 Evaporative Condenser Air Flow Rate [m3/s]
    ,                        !- Speed 3 Rated Evaporative Condenser Pump Power Consumption [W]
    {16},                    !- Speed 4 Gross Rated Total Cooling Capacity [W]
    {8},                     !- Speed 4 Gross Rated Sensible Heat Ratio
    {9},                     !- Speed 4 Gross Rated Cooling COP [W/W]
    {17},                    !- Speed 4 Rated Air Flow Rate [m3/s]
    {6},                     !- Speed 4 2017 Rated Evaporator Fan Power Per Volume Flow Rate [W/(m3/s)]
    {7},                     !- Speed 4 2023 Rated Evaporator Fan Power Per Volume Flow [W/(m3/s)]
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_CapFT,  !- Speed 4 Total Cooling Capacity Function of Temperature Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_CapFF,  !- Speed 4 Total Cooling Capacity Function of Flow Fraction Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_EIRFT,  !- Speed 4 Energy Input Ratio Function of Temperature Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_EIRFFF,  !- Speed 4 Energy Input Ratio Function of Flow Fraction Curve Name
    PSZ-AC_CoolCLennoxStandard10Ton_TGA120S2B_PLR,  !- Speed 4 Part Load Fraction Correlation Curve Name
    ,                        !- Speed 4 Nominal Time for Condensate Removal to Begin [s]
    ,                        !- Speed 4 Ratio of Initial Moisture Evaporation Rate and steady state Latent Capacity [dimensionless]
    ,                        !- Speed 4 Maximum Cycling Rate [cycles/hr]
    ,                        !- Speed 4 Latent Capacity Time Constant [s]
    0.2,                     !- Speed 4 Rated Waste Heat Fraction of Power Input [dimensionless]
    ,                        !- Speed 4 Waste Heat Function of Temperature Curve Name
    0.9,                     !- Speed 4 Evaporative Condenser Effectiveness [dimensionless]
    ,                        !- Speed 4 Evaporative Condenser Air Flow Rate [m3/s]
    ;                        !- Speed 4 Rated Evaporative Condenser Pump Power Consumption [W]";

            return string.Format(
                template,
                coil["Name"].Value,
                coil["Availability Schedule Name"].Value,
                coil["Air Inlet Node Name"].Value,
                coil["Air Outlet Node Name"].Value,
                coil["Crankcase Heater Capacity"].Value,
                coil["Maximum Outdoor Dry-Bulb Temperature for Crankcase Heater Operation"].Value,
                coil["2017 Rated Evaporator Fan Power Per Volume Flow Rate"].Value,
                coil["2023 Rated Evaporator Fan Power Per Volume Flow Rate"].Value,
                shr,
                cop,
                speed1Capacity,
                speed1AirFlowRate,
                speed2Capacity,
                speed2AirFlowRate,
                speed3Capacity,
                speed3AirFlowRate,
                ratedCapacity,
                ratedAirFlowRate);
        }

        public IdfObject ReplaceObjectTypeInList(string listName, string oldObjectType, string oldObjectName, string newObjectType, string newObjectName)
        {
            IEnumerable<IdfObject> allEquipment = Reader[listName];
            foreach (IdfObject equipment in allEquipment)
            {
                for (int i = 0; i < (equipment.Count - 1); i++)
                {
                    Field field = equipment[i];
                    Field nextField = equipment[i + 1];

                    if (field.Value.ToLower() == oldObjectType.ToLower() && nextField.Value.ToLower() == oldObjectName.ToLower())
                    {
                        field.Value = newObjectType;
                        nextField.Value = newObjectName;
                        return equipment;
                    }
                }
            }
            throw new Exception("Could not find any " + listName + " containing reference to " + oldObjectName + " " + oldObjectType);
        }
        public IdfObject FindObject(string objectType, string objectName)
        {
            try
            {
                return Reader[objectType].First(c => c[0] == objectName);
            }
            catch (Exception e)
            {
                throw new Exception(String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }
        private double ConvertWithError(string val, string name, string fieldName)
        {
            try
            {
                return Convert.ToDouble(val);
            }
            catch (FormatException)
            {
                throw new Exception("Cannot replace coil: " + name + ", the " + fieldName + " input needs to be specified.");
            }
        }
    }
}