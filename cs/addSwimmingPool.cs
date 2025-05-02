/*
Add the "SwimmingPool:Indoor" object into the idf file and connect its nodes to the specified HW Loop.

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

        public IdfObject FindObject(string objectType, string objectName)
        {
            try
            {
                return Reader[objectType].First(c => c[0] == objectName);
            }
            catch (Exception e)
            {
                throw new MissingFieldException(String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        public override void BeforeEnergySimulation()
        {
            Reader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath
            );

            AddSwimmingPool("Pool 1", "BLOCK1:SWIMMINGPOOL_GroundFloor_6_0_0", "HW Loop", 1.5, "PoolSetpointTempSched1");

            LoadSharedSchedules();
            LoadSwimmingPoolOutputs();

            Reader.Save();
        }

        private void AddSwimmingPool(string swimmingPoolName, string floorSurfaceName, string hwLoopName, double averageDepth, string setpointScheduleName)
        {
            string branchName = swimmingPoolName + " Branch";
            string swimmingPoolObjects = GetSwimmingPoolContent(swimmingPoolName, branchName, floorSurfaceName, averageDepth, setpointScheduleName);
            Reader.Load(swimmingPoolObjects);

            ConnectSwimmingPoolBranches(branchName, hwLoopName);
        }

        private void ConnectSwimmingPoolBranches(string branchName, string hwLoopName)
        {
            IdfObject hwBranchList = FindObject("BranchList", hwLoopName + " Demand Side Branches");
            IdfObject splitter = FindObject("Connector:Splitter", hwLoopName + " Demand Splitter");
            IdfObject mixer = FindObject("Connector:Mixer", hwLoopName + " Demand Mixer");

            hwBranchList.InsertField(hwBranchList.Count - 2, branchName);
            splitter.InsertField(splitter.Count - 2, branchName);
            mixer.InsertField(mixer.Count - 2, branchName);
        }

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

            return String.Format(template,
                swimmingPoolName,
                floorSurfaceName,
                branchName,
                poolDepth.ToString(System.Globalization.CultureInfo.InvariantCulture),
                setpointScheduleName);
        }

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