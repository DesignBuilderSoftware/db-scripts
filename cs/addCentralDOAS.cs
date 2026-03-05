/*
Central Dedicated Outdoor Air System (DOAS) Injection Script

This DesignBuilder C# script adds a central Dedicated Outdoor Air System (DOAS) that delivers outdoor air to one or more existing air loops.

Purpose
The script adds a central DOAS which includes:
- A chilled water cooling coil (Coil:Cooling:Water)
- A hot water heating coil (Coil:Heating:Water)
- An optional heat recovery heat exchanger (HeatExchanger:AirToAir:SensibleAndLatent)

How to Use

- Configure the DOAS in BeforeEnergySimulation() by setting:
  - targetAirLoopNames: List of air loop names that receive outdoor air from the DOAS
  - hotWaterPlantLoopName: Name of the HW plant loop to serve the DOAS heating coil
  - chilledWaterPlantLoopName: Name of the CHW plant loop to serve the DOAS cooling coil
  - doasName: Base name used to create all DOAS object names
  - doasSupplyAirTempC: Constant supply air temperature setpoint [°C] (used for DOAS supply temp schedule)
  - enableHeatRecovery: true = HX active, false = HX forced off via schedule

Prerequisites

A) Plant loop objects must exist for BOTH the HW and CHW loops referenced in the configuration. The script expects:
   - BranchList named: "<Loop Name> Demand Side Branches"
   - Connector:Splitter named: "<Loop Name> Demand Splitter"
   - Connector:Mixer named: "<Loop Name> Demand Mixer"

B) Child air loop names must match the names in the IDF.
   IMPORTANT: EnergyPlus 9.4 requires air loop names referenced by AirLoopHVAC:DedicatedOutdoorAirSystem to be ALL CAPS.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class AddCentralDoasToAirLoops : ScriptBase, IScript
    {
        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath
            );

            DoasIdfHandler doasIdfHandler = new DoasIdfHandler(idfReader);

            // ----------------------------
            // USER CONFIGURATION SECTION
            // ----------------------------

            // Air loops connected to the central DOAS (E+ 9.4 requires these names to be ALL CAPS)
            List<string> targetAirLoopNames = new List<string> { "AIRLOOP1" };

            // Plant loop names that serve the DOAS coils (must exist in the base model)
            string chilledWaterPlantLoopName = "CHW Loop";
            string hotWaterPlantLoopName = "HW Loop";

            // DOAS base name used to build all object names
            string doasName = "DOAS1";

            // DOAS supply temperature setpoint [°C] (used in a constant Schedule:Compact)
            double doasSupplyAirTempC = 17.5;

            // Enable/disable heat recovery HX
            bool enableHeatRecovery = true;

            // Build the DOAS specification (single object controls all inserted IDF content)
            DoasSpecs doasSpecs = new DoasSpecs(
               doasName,
               targetAirLoopNames,
               hotWaterPlantLoopName,
               chilledWaterPlantLoopName,
               enableHeatRecovery,
               doasSupplyAirTempC
            );

            // Optional: display the configuration summary (comment out to disable message box pop-up)
            MessageBox.Show(doasSpecs.GetInfo());

            // Inject and connect the DOAS into the IDF (loads objects, wires plant branches, optional HX disable, saves IDF)
            doasIdfHandler.LoadDoas(doasSpecs);
        }
    }

    public class DoasSpecs
    {
        public string Name;
        public List<string> ChildAirLoops;
        public string HotWaterPlantLoopName;
        public string ChilledWaterPlantLoopName;
        public bool EnableHeatRecovery;
        public double SupplyTemperatureC;

        public DoasSpecs() { }

        public DoasSpecs(
           string name,
           List<string> childAirLoops,
           string hotWaterPlantLoopName,
           string chilledWaterPlantLoopName,
           bool enableHeatRecovery,
           double supplyTemperatureC)
        {
            Name = name;
            ChildAirLoops = childAirLoops;
            HotWaterPlantLoopName = hotWaterPlantLoopName;
            ChilledWaterPlantLoopName = chilledWaterPlantLoopName;
            EnableHeatRecovery = enableHeatRecovery;
            SupplyTemperatureC = supplyTemperatureC;
        }

        public string HwDemandBranchName { get { return this.Name + " DOAS Heating Coil HW Loop Demand Side Branch"; } }
        public string ChwDemandBranchName { get { return this.Name + " DOAS Cooling Coil CHW Loop Demand Side Branch"; } }
        public string HeatRecoveryHxName { get { return this.Name + " DOAS Heat Recovery Device"; } }
        public string AlwaysOffScheduleName { get { return this.Name + " ALWAYS_OFF"; } }

        public string GetInfo()
        {
            string childLoopNames = String.Join("\n - ", this.ChildAirLoops);
            string text = @"DOAS: {0}
HW loop: {1}
CHW loop: {2}
Heat recovery enabled: {3}
Supply temperature [C]: {4}
Child air loops:
- {5}";
            return string.Format(
               text,
               this.Name,
               this.HotWaterPlantLoopName,
               this.ChilledWaterPlantLoopName,
               this.EnableHeatRecovery,
               this.SupplyTemperatureC,
               childLoopNames);
        }

        public string BuildIdfObjectsText()
        {
            string airLoops = String.Join(",\n", this.ChildAirLoops);
            int airLoopCount = this.ChildAirLoops.Count;
            string airLoopInlets = String.Join(",\n", this.ChildAirLoops.Select(x => x + " AHU Outdoor Air Inlet"));
            string airLoopOutlets = String.Join(",\n", this.ChildAirLoops.Select(x => x + " AHU Relief Air Outlet"));

            string idfObjects = @"
!-   ===========  ALL OBJECTS IN CLASS: SCHEDULE:COMPACT ===========

Schedule:Compact,
   {0} DOAS Supply Air Temp Sch,                                  !- Name
   Temperature,                                                   !- Schedule Type Limits Name
   Through: 12/31,                                                !- Field 1
   For: AllDays,                                                  !- Field 2
   Until: 24:00,                                                  !- Field 3
   {5};                                                           !- Field 4

Schedule:Compact,
   {0} ALWAYS_ON,                                                 !- Name
   On/Off,                                                        !- Schedule Type Limits Name
   Through: 12/31,                                                !- Field 1
   For: AllDays,                                                  !- Field 2
   Until: 24:00,                                                  !- Field 3
   1;                                                             !- Field 4

Schedule:Compact,
   {0} ALWAYS_OFF,                                                !- Name
   On/Off,                                                        !- Schedule Type Limits Name
   Through: 12/31,                                                !- Field 1
   For: AllDays,                                                  !- Field 2
   Until: 24:00,                                                  !- Field 3
   0;                                                             !- Field 4


!-   ===========  ALL OBJECTS IN CLASS: FAN:SYSTEMMODEL ===========

Fan:SystemModel,
   {0} DOAS OA Supply Fan,                                        !- Name
   {0} ALWAYS_ON,                                                 !- Availability Schedule Name
   {0} DOAS Heating Coil Air Outlet Node,                         !- Air Inlet Node Name
   {0} AirLoopSplitterInlet,                                      !- Air Outlet Node Name
   Autosize,                                                      !- Design Maximum Air Flow Rate m3/s
   Discrete,                                                      !- Speed Control Method
   0.25,                                                          !- Electric Power Minimum Flow Rate Fraction
   600.0,                                                         !- Design Pressure Rise Pa
   0.9,                                                           !- Motor Efficiency
   1.0,                                                           !- Motor In Air Stream Fraction
   Autosize,                                                      !- Design Electric Power Consumption W
   TotalEfficiencyAndPressure,                                    !- Design Power Sizing Method
   ,                                                              !- Electric Power Per Unit Flow Rate W/(m3/s)
   ,                                                              !- Electric Power Per Unit Flow Rate Per Unit Pressure W/((m3/s)-Pa)
   0.7,                                                           !- Fan Total Efficiency
   ,                                                              !- Electric Power Function of Flow Fraction Curve Name
   ,                                                              !- Night Ventilation Mode Pressure Rise Pa
   ,                                                              !- Night Ventilation Mode Flow Fraction
   ,                                                              !- Motor Loss Zone Name
   ,                                                              !- Motor Loss Radiative Fraction
   General;                                                       !- End-Use Subcategory


!-   ===========  ALL OBJECTS IN CLASS: HEATEXCHANGER:AIRTOAIR:SENSIBLEANDLATENT ===========

HeatExchanger:AirToAir:SensibleAndLatent,
   {0} DOAS Heat Recovery Device,                                 !- Name
   {0} ALWAYS_ON,                                                 !- Availability Schedule Name
   autosize,                                                      !- Nominal Supply Air Flow Rate m3/s
   0.750,                                                         !- Sensible Effectiveness at 100% Heating Air Flow dimensionless
   0.000,                                                         !- Latent Effectiveness at 100% Heating Air Flow dimensionless
   0.750,                                                         !- Sensible Effectiveness at 75% Heating Air Flow dimensionless
   0.000,                                                         !- Latent Effectiveness at 75% Heating Air Flow dimensionless
   0.750,                                                         !- Sensible Effectiveness at 100% Cooling Air Flow dimensionless
   0.000,                                                         !- Latent Effectiveness at 100% Cooling Air Flow dimensionless
   0.750,                                                         !- Sensible Effectiveness at 75% Cooling Air Flow dimensionless
   0.000,                                                         !- Latent Effectiveness at 75% Cooling Air Flow dimensionless
   {0} Outside Air Inlet Node 1,                                  !- Supply Air Inlet Node Name
   {0} DOAS Heat Recovery Device Supply Outlet,                   !- Supply Air Outlet Node Name
   {0} AirLoopDOASMixerOutlet,                                    !- Exhaust Air Inlet Node Name
   {0} DOAS Heat Recovery Device Relief Outlet,                   !- Exhaust Air Outlet Node Name
   0.000,                                                         !- Nominal Electric Power W
   No,                                                            !- Supply Air Outlet Temperature Control
   Plate,                                                         !- Heat Exchanger Type
   None,                                                          !- Frost Control Type
   1.70,                                                          !- Threshold Temperature C
   0.167,                                                         !- Initial Defrost Time Fraction dimensionless
   0.0240,                                                        !- Rate of Defrost Time Fraction Increase 1/K
   Yes;                                                           !- Economizer Lockout


!-   ===========  ALL OBJECTS IN CLASS: AIRLOOPHVAC:OUTDOORAIRSYSTEM:EQUIPMENTLIST ===========

AirLoopHVAC:OutdoorAirSystem:EquipmentList,
   {0} OA Sys Equipment,                                          !- Name
   HeatExchanger:AirToAir:SensibleAndLatent,                      !- Component 1 Object Type
   {0} DOAS Heat Recovery Device,                                 !- Component 1 Name
   Coil:Cooling:Water,                                            !- Component 2 Object Type
   {0} DOAS CHW Cooling Coil,                                     !- Component 2 Name
   Coil:Heating:Water,                                            !- Component 3 Object Type
   {0} DOAS HW Heating Coil,                                      !- Component 3 Name
   Fan:SystemModel,                                               !- Component 4 Object Type
   {0} DOAS OA Supply Fan;                                        !- Component 4 Name


!-   ===========  ALL OBJECTS IN CLASS: AIRLOOPHVAC:OUTDOORAIRSYSTEM ===========

AirLoopHVAC:OutdoorAirSystem,
   {0} AirLoop DOAS OA system,                                    !- Name
   {0} OA Sys Controllers,                                        !- Controller List Name
   {0} OA Sys Equipment,                                          !- Outdoor Air Equipment List Name
   {0} OA Sys Avail List;                                         !- Availability Manager List Name


!-   ===========  ALL OBJECTS IN CLASS: AIRLOOPHVAC:DEDICATEDOUTDOORAIRSYSTEM ===========

AirLoopHVAC:DedicatedOutdoorAirSystem,
   {0},                                                           !- Name
   {0} AirLoop DOAS OA system,                                    !- AirLoopHVAC:OutdoorAirSystem Name
   {0} ALWAYS_ON,                                                 !- Availability Schedule Name
   {0} AirLoopDOASMixer,                                          !- AirLoopHVAC:Mixer Name
   {0} AirLoopDOASSplitter,                                       !- AirLoopHVAC:Splitter Name
   {5},                                                           !- Preheat Design Temperature C
   0.004,                                                         !- Preheat Design Humidity Ratio kgWater/kgDryAir
   {5},                                                           !- Precool Design Temperature C
   0.008,                                                         !- Precool Design Humidity Ratio kgWater/kgDryAir
   {1},                                                           !- Number of AirLoopHVAC
   {2};                                                           !- Air loop names (E+9.4 expects ALL CAPS)


!-   ===========  ALL OBJECTS IN CLASS: AIRLOOPHVAC:MIXER ===========

AirLoopHVAC:Mixer,
   {0} AirLoopDOASMixer,                                          !- Name
   {0} AirLoopDOASMixerOutlet,                                    !- Outlet Node Name
   {3};


!-   ===========  ALL OBJECTS IN CLASS: AIRLOOPHVAC:SPLITTER ===========

AirLoopHVAC:Splitter,
   {0} AirLoopDOASSplitter,                                       !- Name
   {0} AirLoopDOASSplitterInlet,                                  !- Inlet Node Name
   {4};


!-   ===========  ALL OBJECTS IN CLASS: OUTDOORAIR:NODELIST ===========

OutdoorAir:NodeList,
   {0} OutsideAirInletNodes,                                      !- Node or NodeList Name 1
   {0} Outside Air Inlet Node 1;                                  !- Node or NodeList Name 2


!-   ===========  ALL OBJECTS IN CLASS: AVAILABILITYMANAGER:SCHEDULED ===========

AvailabilityManager:Scheduled,
   {0} OA Sys Avail,                                              !- Name
   {0} ALWAYS_ON;                                                 !- Schedule Name


!-   ===========  ALL OBJECTS IN CLASS: AVAILABILITYMANAGERASSIGNMENTLIST ===========

AvailabilityManagerAssignmentList,
   {0} OA Sys Avail List,                                         !- Name
   AvailabilityManager:Scheduled,                                 !- Availability Manager 1 Object Type
   {0} OA Sys Avail;                                              !- Availability Manager 1 Name


!-   ===========  ALL OBJECTS IN CLASS: SETPOINTMANAGER:SCHEDULED ===========

SetpointManager:Scheduled,
   {0} CHW Coil SPM,                                              !- Name
   Temperature,                                                   !- Control Variable
   {0} DOAS Supply Air Temp Sch,                                  !- Schedule Name
   {0} DOAS Cooling Coil Air Outlet Node;                         !- Setpoint Node or NodeList Name

SetpointManager:Scheduled,
   {0} HW Coil SPM,                                               !- Name
   Temperature,                                                   !- Control Variable
   {0} DOAS Supply Air Temp Sch,                                  !- Schedule Name
   {0} DOAS Heating Coil Air Outlet Node;                         !- Setpoint Node or NodeList Name

Coil:Cooling:Water,
  {0} DOAS CHW Cooling Coil,                                      !- Component name
  {0} ALWAYS_ON,                                                  !- Availability schedule
  autosize,                                                       !- Design Water Volume Flow Rate of Coil (m3/s)
  autosize,                                                       !- Design Air Flow Rate of Coil (m3/s)
  autosize,                                                       !- Design Inlet Water Temperature (C)
  autosize,                                                       !- Design Inlet Air Temperature (C)
  autosize,                                                       !- Design Outlet Air Temperature (C)
  autosize,                                                       !- Design Inlet Air Humidity Ratio
  autosize,                                                       !- Design Outlet Air Humidity Ratio
  {0} DOAS Cooling Coil Water Inlet Node,                         !- Water inlet node name
  {0} DOAS Cooling Coil Water Outlet Node,                        !- Water outlet node name
  {0} DOAS Heat Recovery Device Supply Outlet,                    !- Air inlet node name
  {0} DOAS Cooling Coil Air Outlet Node,                          !- Air outlet node name
  SimpleAnalysis,                                                 !- Coil Analysis Type
  CrossFlow,                                                      !- Heat Exchanger Configuration
  ;                                                               !- Water Storage Tank for Condensate Collection

Coil:Heating:Water,
  {0} DOAS HW Heating Coil,                                       !- Component name
  {0} ALWAYS_ON,                                                  !- Availability schedule
  autosize,                                                       !- U-factor times area value of coil (W/K)
  autosize,                                                       !- Max water flow rate of coil (m3/s)
  {0} DOAS Heating Coil Water Inlet Node,                         !- Water inlet node name
  {0} DOAS Heating Coil Water Outlet Node,                        !- Water outlet node name
  {0} DOAS Cooling Coil Air Outlet Node,                          !- Air inlet node name
  {0} DOAS Heating Coil Air Outlet Node,                          !- Air outlet node name
  UFactorTimesAreaAndDesignWaterFlowRate,                         !- Coil performance input method
  autosize,                                                       !- Rated capacity (W)
  80.0,                                                           !- Rated inlet water temperature (C)
  16.0,                                                           !- Rated inlet air temperature (C)
  70.0,                                                           !- Rated outlet water temperature (C)
  35.0,                                                           !- Rated outlet air temperature (C)
  0.50;                                                           !- Rated ratio for air and water convection

Branch,
  {6},                                                            !- Branch name
  ,                                                               !- Pressure drop curve name
  Coil:Cooling:Water,                                             !- Component 1 object type
  {0} DOAS CHW Cooling Coil,                                      !- Component 1 name
  {0} DOAS Cooling Coil Water Inlet Node,                         !- Component 1 inlet node name
  {0} DOAS Cooling Coil Water Outlet Node;                        !- Component 1 outlet node name

Branch,
  {7},                                                            !- Branch name
  ,                                                               !- Pressure drop curve name
  Coil:Heating:Water,                                             !- Component 1 object type
  {0} DOAS HW Heating Coil,                                       !- Component 1 name
  {0} DOAS Heating Coil Water Inlet Node,                         !- Component 1 inlet node name
  {0} DOAS Heating Coil Water Outlet Node;                        !- Component 1 outlet node name

Controller:WaterCoil,
  {0} DOAS Cooling Coil Controller,                               !- Controller name
  Temperature,                                                    !- Control variable
  Reverse,                                                        !- Control action
  Flow,                                                           !- Actuator variable
  {0} DOAS Cooling Coil Air Outlet Node,                          !- Sensor node name
  {0} DOAS Cooling Coil Water Inlet Node,                         !- Actuator node name
  autosize,                                                       !- Controller convergence tolerance
  autosize,                                                       !- Maximum actuated flow (m3/s)
  0.000000;                                                       !- Minimum actuated flow (m3/s)

Controller:WaterCoil,
  {0} DOAS Heating Coil Controller,                               !- Controller name
  Temperature,                                                    !- Control variable
  Normal,                                                         !- Control action
  Flow,                                                           !- Actuator variable
  {0} DOAS Heating Coil Air Outlet Node,                          !- Sensor node name
  {0} DOAS Heating Coil Water Inlet Node,                         !- Actuator node name
  autosize,                                                       !- Controller convergence tolerance
  autosize,                                                       !- Maximum actuated flow (m3/s)
  0.000000;                                                       !- Minimum actuated flow (m3/s)

AirLoopHVAC:ControllerList,
  {0} OA Sys Controllers,
  Controller:WaterCoil,
  {0} DOAS Cooling Coil Controller,
  Controller:WaterCoil,
  {0} DOAS Heating Coil Controller;";

            return String.Format(
               idfObjects,
               this.Name,                    
               airLoopCount,                 
               airLoops,                     
               airLoopOutlets,               
               airLoopInlets,                
               this.SupplyTemperatureC,      
               this.ChwDemandBranchName,     
               this.HwDemandBranchName       
            );
        }
    }

    public class DoasIdfHandler
    {
        public IdfReader Idf;

        public DoasIdfHandler() { }

        public DoasIdfHandler(IdfReader idfReader)
        {
            Idf = idfReader;
        }

        public IdfObject FindObject(string objectType, string objectName)
        {
            try
            {
                return this.Idf[objectType].First(c => c[0] == objectName);
            }
            catch (Exception)
            {
                throw new Exception(String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        private void AddDemandSideBranchToPlantLoop(string plantLoopName, string branchName)
        {
            IdfObject branchList = FindObject("BranchList", plantLoopName + " Demand Side Branches");
            branchList.InsertField(branchList.Count - 1, branchName);

            IdfObject splitter = FindObject("Connector:Splitter", plantLoopName + " Demand Splitter");
            splitter.InsertField(splitter.Count - 1, branchName);

            IdfObject mixer = FindObject("Connector:Mixer", plantLoopName + " Demand Mixer");
            mixer.InsertField(mixer.Count - 1, branchName);
        }

        public void LoadDoas(DoasSpecs doasSpecs)
        {
            // Load all DOAS-related IDF objects (schedules, OA system, DOAS object, coils, controllers, branches)
            string doasIdfObjectsText = doasSpecs.BuildIdfObjectsText();
            this.Idf.Load(doasIdfObjectsText);

            // Connect the DOAS heating/cooling coil branches to the HW/CHW plant loop demand sides
            AddDemandSideBranchToPlantLoop(doasSpecs.HotWaterPlantLoopName, doasSpecs.HwDemandBranchName);
            AddDemandSideBranchToPlantLoop(doasSpecs.ChilledWaterPlantLoopName, doasSpecs.ChwDemandBranchName);

            // If heat recovery is disabled, force the HX off via its availability schedule
            if (!doasSpecs.EnableHeatRecovery)
            {
                IdfObject hx = FindObject("HeatExchanger:AirToAir:SensibleAndLatent", doasSpecs.HeatRecoveryHxName);
                hx[1].Value = doasSpecs.AlwaysOffScheduleName;
            }

            // Save the modified IDF
            this.Idf.Save();
        }
    }
}