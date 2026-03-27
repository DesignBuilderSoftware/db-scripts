/*
Indoor Swimming Pool (EnergyPlus) – IDF Injection + HW Loop Demand Connection Script

Purpose:
This DesignBuilder C# script adds an EnergyPlus `SwimmingPool:Indoor` object to the exported IDF and connects it to a specified Hot Water (HW) plant loop demand side.
It runs automatically before the EnergyPlus simulation starts and modifies the IDF in-place.

Main Steps:
1) Create a new demand-side `Branch` containing the `SwimmingPool:Indoor` component and its inlet/outlet nodes.
2) Insert that branch into the target HW loop demand side.
3) Load shared schedules used by the pool object (activity, cover, make-up water, occupancy, setpoint).
4) Add Output:Variable requests relevant to Indoor Swimming Pool reporting.
5) Save the modified IDF.

How to Use:

Configuration
Edit the `AddSwimmingPool(...)` call(s) inside `BeforeEnergySimulation()`:
- swimmingPoolName: Unique pool name used for the SwimmingPool:Indoor object and to generate node names.
- floorSurfaceName: Name of the floor surface that represents the pool surface (must match an existing Surface name in the IDF).
- hwLoopName: Base name of the HW PlantLoop used to locate demand-side objects.
- averageDepth: Pool average depth in meters.
- setpointScheduleName: Name of a temperature schedule (must exist or be created by this script).

Prerequisites / Placeholders
The base model / exported IDF must already contain:
- A How Water Loop (referenced as "HW Loop" in the example case).
- A surface with name exactly matching your model's surface in which the pool will be located
    (referenced as "BLOCK1:SWIMMINGPOOL_GroundFloor_6_0_0" in the example case).
    NOTE: The surface name can be found Miscellaneous model data tab under the Name in last EnergyPlus calculation.
- A setpoint schedule for the pool water temperature (referenced as "PoolSetpointTempSched1" in the example case)

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
        private IdfReader Reader;

        public IdfObject FindObject(string objectType, string objectName)
        {
            try
            {
                return Reader[objectType].First(c => c[0] == objectName);
            }
            catch (Exception)
            {
                throw new MissingFieldException(
                    String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        public override void BeforeEnergySimulation()
        {
            Reader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath
            );

            // ----------------------------
            // USER CONFIGURATION SECTION
            // ----------------------------
            // Add one or more pools by calling AddSwimmingPool() with the required names and parameters.
            AddSwimmingPool(
                "Pool 1", // Pool object name
                "BLOCK1:SWIMMINGPOOL_GroundFloor_6_0_0", // Surface name (as defined in the model)
                "HW Loop", // Hot Water loop name (as defined in the model)
                1.5,  // Pool average depth, in meters
                "PoolSetpointTempSched1" // Pool water temperature schedule (as defined in the model or included in the script)
            );

            LoadSharedSchedules();
            LoadSwimmingPoolOutputs();

            Reader.Save();
        }

        private void AddSwimmingPool(string swimmingPoolName, string floorSurfaceName, string hwLoopName, double averageDepth, string setpointScheduleName)
        {
            string branchName = swimmingPoolName + " Branch";

            // Create the Branch + SwimmingPool:Indoor IDF text and load it into the model.
            string swimmingPoolObjects = GetSwimmingPoolContent(swimmingPoolName, branchName, floorSurfaceName, averageDepth, setpointScheduleName);
            Reader.Load(swimmingPoolObjects);

            // Connect the new branch into the HW PlantLoop demand BranchList / Splitter / Mixer.
            ConnectSwimmingPoolBranches(branchName, hwLoopName);
        }

        // Inserts the pool branch into the HW loop demand-side BranchList and the corresponding splitter/mixer.
        private void ConnectSwimmingPoolBranches(string branchName, string hwLoopName)
        {
            IdfObject hwBranchList = FindObject("BranchList", hwLoopName + " Demand Side Branches");
            IdfObject splitter = FindObject("Connector:Splitter", hwLoopName + " Demand Splitter");
            IdfObject mixer = FindObject("Connector:Mixer", hwLoopName + " Demand Mixer");

            // Insert the new demand branch name into the demand BranchList and connector nodes list.
            hwBranchList.InsertField(hwBranchList.Count - 2, branchName);
            splitter.InsertField(splitter.Count - 2, branchName);
            mixer.InsertField(mixer.Count - 2, branchName);
        }

        // Returns the IDF text for yhe SwimmingPool:Indoor and the respective demand-side branch.
        private string GetSwimmingPoolContent(string swimmingPoolName, string branchName, string floorSurfaceName, double poolDepth, string setpointScheduleName)
        {
            string template = @"
Branch,
    {2},                     !- Name
    ,                        !- Pressure Drop Curve Name
    SwimmingPool:Indoor,     !- Component 1 Object Type
    {0},                     !- Component 1 Name
    {0}  Water Inlet Node,   !- Component 1 Inlet Node Name
    {0}  Water Outlet Node;  !- Component 1 Outlet Node Name

SwimmingPool:Indoor,
  {0},                     !- Name
  {1},                     !- Surface Name
  {3},                     !- Average Depth m
  PoolActivitySched,       !- Activity Factor Schedule Name
  MakeUpWaterSched,        !- Make-up Water Supply Schedule Name
  PoolCoverSched,          !- Cover Schedule Name
  0.8,                     !- Cover Evaporation Factor
  0.2,                     !- Cover Convection Factor
  0.9,                     !- Cover Short-Wavelength Radiation Factor
  0.5,                     !- Cover Long-Wavelength Radiation Factor
  {0}  Water Inlet Node,   !- Pool Water Inlet Node
  {0}  Water Outlet Node,  !- Pool Water Outlet Node
  0.1,                     !- Pool Heating System Maximum Water Flow Rate m3/s
  0.6,                     !- Pool Miscellaneous Equipment Power W/(m3/s)
  {4},                    !- Setpoint Temperature Schedule
  10,                      !- Maximum Number of People
  PoolOccupancySched,      !- People Schedule
  PoolOccHeatGainSched;    !- People Heat Gain Schedule";

            return String.Format(
                template,
                swimmingPoolName,
                floorSurfaceName,
                branchName,
                poolDepth.ToString(System.Globalization.CultureInfo.InvariantCulture),
                setpointScheduleName
            );
        }

        // Loads schedules referenced by the SwimmingPool:Indoor object(s) created by this script.
        // NOTE: If you change schedule names in GetSwimmingPoolContent(), update them here as well.
        private void LoadSharedSchedules()
        {
            string template = @"
Schedule:Compact,
  PoolActivitySched,
  Fraction,
  Through: 12/31,
  For: WeekDays SummerDesignDay,
  Until: 6:00,0.1,
  Until: 20:00,0.5,
  Until: 24:00,0.1,
  For: AllOtherDays,
  Until: 24:00,0.1;

Schedule:Compact,
  MakeUpWaterSched,
  Any Number,
  Through: 12/31,
  For: AllDays,
  Until: 24:00,16.67;

Schedule:Compact,
  PoolCoverSched,
  Fraction,
  Through: 12/31,
  For: WeekDays SummerDesignDay,
  Until: 6:00,0.5,
  Until: 20:00,0.0,
  Until: 24:00,0.5,
  For: AllOtherDays,
  Until: 24:00,1.0;

Schedule:Compact,
  PoolSetpointTempSched1,
  Any Number,
  Through: 12/31,
  For: AllDays,
  Until: 24:00,27.0;

Schedule:Compact,
  PoolSetpointTempSched2,
  Any Number,
  Through: 12/31,
  For: AllDays,
  Until: 24:00,29.0;

Schedule:Compact,
  PoolOccupancySched,
  Fraction,
  Through: 12/31,
  For: WeekDays SummerDesignDay,
  Until: 6:00,0.0,
  Until: 9:00,1.0,
  Until: 11:00,0.5,
  Until: 13:00,1.0,
  Until: 16:00,0.5,
  Until: 20:00,1.0,
  Until: 24:00,0.0,
  For: AllOtherDays,
  Until: 24:00,0.0;

Schedule:Compact,
  PoolOccHeatGainSched,
  Any Number,
  Through: 12/31,
  For: AllDays,
  Until: 24:00,300.0;";
            Reader.Load(template);
        }

        // Requests common output variables for Indoor Swimming Pool reporting at hourly frequency.
        private void LoadSwimmingPoolOutputs()
        {
            string template = @"
Output:Variable,*,Indoor Pool Makeup Water Rate,hourly;
Output:Variable,*,Indoor Pool Makeup Water Volume,hourly;
Output:Variable,*,Indoor Pool Makeup Water Temperature,hourly;
Output:Variable,*,Indoor Pool Water Temperature,hourly;
Output:Variable,*,Indoor Pool Inlet Water Temperature,hourly;
Output:Variable,*,Indoor Pool Inlet Water Mass Flow Rate,hourly;
Output:Variable,*,Indoor Pool Miscellaneous Equipment Power,hourly;
Output:Variable,*,Indoor Pool Miscellaneous Equipment Energy,hourly;
Output:Variable,*,Indoor Pool Water Heating Rate,hourly;
Output:Variable,*,Indoor Pool Water Heating Energy,hourly;
Output:Variable,*,Indoor Pool Radiant to Convection by Cover,hourly;
Output:Variable,*,Indoor Pool People Heat Gain,hourly;
Output:Variable,*,Indoor Pool Current Activity Factor,hourly;
Output:Variable,*,Indoor Pool Current Cover Factor,hourly;
Output:Variable,*,Indoor Pool Saturation Pressure at Pool Temperature,hourly;
Output:Variable,*,Indoor Pool Partial Pressure of Water Vapor in Air,hourly;
Output:Variable,*,Indoor Pool Current Cover Evaporation Factor,hourly;
Output:Variable,*,Indoor Pool Current Cover Convective Factor,hourly;
Output:Variable,*,Indoor Pool Current Cover SW Radiation Factor,hourly;
Output:Variable,*,Indoor Pool Current Cover LW Radiation Factor,hourly;
Output:Variable,*,Indoor Pool Evaporative Heat Loss Rate,hourly;
Output:Variable,*,Indoor Pool Evaporative Heat Loss Energy,hourly;";
            Reader.Load(template);
        }
    }
}