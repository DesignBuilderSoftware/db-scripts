/*
This script adds a cental deficated outdoor air system unit to deliver outdoor air to specified air loops.

The unit includes hot and chilled water coils to pre-treat the air and an optional heat exchanger.

The system is specified via the 'DoasSpecs' object.
Example specification is present in the 'BeforeEnergySimulation' hookpoint below.

Note that E+9.4 requires air loop names to be in ALL CAPS.
*/

using System;
using System.Runtime;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
   public class IdfFindAndReplace : ScriptBase, IScript
   {
       public override void BeforeEnergySimulation()
       {
           IdfReader idfReader = new IdfReader(
               ApiEnvironment.EnergyPlusInputIdfPath,
               ApiEnvironment.EnergyPlusInputIddPath
           );

           DoasIdfHandler idfHandler = new DoasIdfHandler(idfReader);

           List<string> doas1AirLoops = new List<string> { "AIRLOOP1" }; // specify air loops connected to the central DOAS
           string chwLoopName = "CHW Loop"; // specify chw loop connected to the central DOAS cooling coil
           string hwLoopName = "HW Loop";  // specify hw loop connected to the central DOAS heating coil
           string doasName = "DOAS1";
           double supplyAirTemperature = 17.5; // supply air temperature in degrees of Celsius
           bool includeHX = true;

           DoasSpecs doas1 = new DoasSpecs(doasName, doas1AirLoops, hwLoopName, chwLoopName, includeHX, supplyAirTemperature);

           // display doas specification (comment out to disable the message box pop-up)
           MessageBox.Show(doas1.GetInfo());

           idfHandler.LoadDoas(doas1);
       }
   }

   public class DoasSpecs
   {
       public string Name;
       public List<string> ChildAirLoops;
       public string HwLoopName;
       public string ChwLoopName;
       public bool IncludeHX;
       public double SupplyTemperature;

       public DoasSpecs() {}

       public DoasSpecs(string name, List<string> childAirLoops, string hwLoopName, string chwLoopName, bool includeHX, double supplyTemperature)
       {
           Name = name;
           ChildAirLoops = childAirLoops;
           HwLoopName = hwLoopName;
           ChwLoopName = chwLoopName;
           IncludeHX = includeHX;
           SupplyTemperature = supplyTemperature;
       }

       public string HwBranchName { get { return this.Name + " DOAS Heating Coil HW Loop Demand Side Branch"; } }
       public string ChwBranchName { get { return this.Name + " DOAS Cooling Coil CHW Loop Demand Side Branch"; } }
       public string FanName { get { return this.Name + "DOAS OA Supply Fan"; } }
       public string HxName { get { return this.Name + " DOAS Heat Recovery Device"; } }
       public string OffScheduleName { get { return this.Name + " ALWAYS_OFF"; } }


       public string GetInfo()
       {
           string childLoopNames = String.Join("\n - ", this.ChildAirLoops);
           string text = @"DOAS: {0}
HW loop: {1}
CHW loop: {2}
HX included: {3}
Supply temperature: {4}
Child air loops:
- {5}";
           return string.Format(text, this.Name, this.HwLoopName, this.ChwLoopName, this.IncludeHX, this.SupplyTemperature, childLoopNames);
       }

       public string GetIDFObjects()
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
   {0}DOAS OA Supply Fan,                                             !- Name
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
   Coil:Heating:Water,                                            !- Component 2 Object Type
   {0} DOAS HW Heating Coil,                                      !- Component 2 Name
   Fan:SystemModel,                                               !- Component 3 Object Type
   {0}DOAS OA Supply Fan;                                             !- Component 3 Name


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
   {2};                                                           ! Air loop names (must be ALL CAPS)


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
   {0} Always_ON;                                                 !- Schedule Name


!-   ===========  ALL OBJECTS IN CLASS: AVAILABILITYMANAGERASSIGNMENTLIST ===========

AvailabilityManagerAssignmentList,
   {0} OA Sys Avail List,                                         !- Name
   AvailabilityManager:Scheduled,                                 !- Availability Manager 1 Object Type
   {0} OA SysAvail;                                               !- Availability Manager 1 Name


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
  {0} DOAS CHW Cooling Coil,                                      ! - Component name
  {0} ALWAYS_ON,                                                  ! - Availability schedule
  autosize,                                                       ! - Design Water Volume Flow Rate of Coil (m3/s)
  autosize,                                                       ! - Design Air Flow Rate of Coil (m3/s)
  autosize,                                                       ! - Design Inlet Water Temperature (C)
  autosize,                                                       ! - Design Inlet Air Temperature (C)
  autosize,                                                       ! - Design Outlet Air Temperature (C)
  autosize,                                                       ! - Design Inlet Air Humidity Ratio
  autosize,                                                       ! - Design Outlet Air Humidity Ratio
  {0} DOAS Cooling Coil Water Inlet Node,                         ! - Water inlet node name
  {0} DOAS Cooling Coil Water Outlet Node,                        ! - Water outlet node name
  {0} DOAS Heat Recovery Device Supply Outlet,                    ! - Air inlet node name
  {0} DOAS Cooling Coil Air Outlet Node,                          ! - Air outlet node name
  SimpleAnalysis,                                                 ! - Coil Analysis Type
  CrossFlow,                                                      ! - Heat Exchanger Configuration
  ;                                                               ! - Water Storage Tank for Condensate Collection

 Coil:Heating:Water,
  {0} DOAS HW Heating Coil,                                       ! - Component name
  {0} ALWAYS_ON,                                                  ! - Availability schedule
  autosize,                                                       ! - U-factor times area value of coil (W/K)
  autosize,                                                       ! - Max water flow rate of coil (m3/s)
  {0} DOAS Heating Coil Water Inlet Node,                         ! - Water inlet node name
  {0} DOAS Heating Coil Water Outlet Node,                        ! - Water outlet node name
  {0} DOAS Cooling Coil Air Outlet Node,                          ! - Air inlet node name
  {0} DOAS Heating Coil Air Outlet Node,                          ! - Air outlet node name
  UFactorTimesAreaAndDesignWaterFlowRate,                         ! - Coil performance input method
  autosize,                                                       ! - Rated capacity (W)
  80.0,                                                           ! - Rated inlet water temperature (C)
  16.0,                                                           ! - Rated inlet air temperature (C)
  70.0,                                                           ! - Rated outlet water temperature (C)
  35.0,                                                           ! - Rated outlet air temperature (C)
  0.50;                                                           ! - Rated ratio for air and water convection

Branch,
  {6},                                                            ! - Branch name
  ,                                                               ! - Pressure drop curve name
  Coil:Cooling:Water,                                             ! - Component 1 object type
  {0} DOAS CHW Cooling Coil,                                      ! - Component 1 name
  {0} DOAS Cooling Coil Water Inlet Node,                         ! - Component 1 inlet node name
  {0} DOAS Cooling Coil Water Outlet Node;                        ! - Component 1 outlet node name

Branch,
  {7},                                                            ! - Branch name
  ,                                                               ! - Pressure drop curve name
  Coil:Heating:Water,                                             ! - Component 1 object type
  {0} DOAS HW Heating Coil,                                       ! - Component 1 name
  {0} DOAS Heating Coil Water Inlet Node,                         ! - Component 1 inlet node name
  {0} DOAS Heating Coil Water Outlet Node;                        ! - Component 1 outlet node name

 Controller:WaterCoil,
  {0} DOAS Cooling Coil Controller,                               ! - Controller name
  Temperature,                                                    ! - Control variable
  Reverse,                                                        ! - Control action
  Flow,                                                           ! - Actuator variable
  {0} DOAS Cooling Coil Air Outlet Node,                          ! - Sensor node name
  {0} DOAS Cooling Coil Water Inlet Node,                         ! - Actuator node name
  autosize,                                                       ! - Controller convergence tolerance
  autosize,                                                       ! - Maximum actuated flow (m3/s)
  0.000000;                                                       ! - Minimum actuated flow (m3/s)

Controller:WaterCoil,
  {0} DOAS Heating Coil Controller,                               ! - Controller name
  Temperature,                                                    ! - Control variable
  Normal,                                                         ! - Control action
  Flow,                                                           ! - Actuator variable
  {0} DOAS Heating Coil Air Outlet Node,                          ! - Sensor node name
  {0} DOAS Heating Coil Water Inlet Node,                         ! - Actuator node name
  autosize,                                                       ! - Controller convergence tolerance
  autosize,                                                       ! - Maximum actuated flow (m3/s)
  0.000000;                                                       ! - Minimum actuated flow (m3/s)

AirLoopHVAC:ControllerList,
  {0} OA Sys Controllers,
  Controller:WaterCoil,
  {0} DOAS Cooling Coil Controller,
  Controller:WaterCoil,
  {0} DOAS Heating Coil Controller;";

           return String.Format(idfObjects, this.Name, airLoopCount, airLoops, airLoopOutlets, airLoopInlets, this.SupplyTemperature, this.ChwBranchName, this.HwBranchName);
       }
   }

   public class DoasIdfHandler
   {
       public IdfReader Reader;

       public DoasIdfHandler(){}

       public DoasIdfHandler(IdfReader idfReader)
       {
           Reader = idfReader;
       }

       public IdfObject FindObject(string objectType, string objectName)
       {
           try
           {
               return this.Reader[objectType].First(c => c[0] == objectName);
           }
           catch(Exception e)
           {
               throw new Exception(String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
           }
       }

       private void AddBranch(string loopName, string branchName)
       {
           IdfObject branchList = FindObject("branchList", loopName + " Demand Side Branches");
           branchList.InsertField(branchList.Count - 1, branchName);

           IdfObject splitter = FindObject("Connector:Splitter", loopName + " Demand Splitter");
           splitter.InsertField(splitter.Count - 1, branchName);

           IdfObject mixer = FindObject( "Connector:Mixer", loopName + " Demand Mixer");
           mixer.InsertField(mixer.Count - 1, branchName);
       }

       public void LoadDoas(DoasSpecs doasSpecs)
       {
           string doasIdfObjects = doasSpecs.GetIDFObjects();
           this.Reader.Load(doasIdfObjects);

           AddBranch(doasSpecs.HwLoopName, doasSpecs.HwBranchName);
           AddBranch(doasSpecs.ChwLoopName, doasSpecs.ChwBranchName);

           if (!doasSpecs.IncludeHX)
           {
               IdfObject hx = FindObject("HeatExchanger:AirToAir:SensibleAndLatent", doasSpecs.HxName);
               hx[1].Value = doasSpecs.OffScheduleName;
           }

           this.Reader.Save();
       }
   }
}
