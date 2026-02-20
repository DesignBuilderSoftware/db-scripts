/*
Add EvaporativeCooler:Direct:ResearchSpecial to air loop return air stream.

The script adds the component to air loops specified in "airLoopNames" array.
Cooler parameters and setpoint can be set in the object boilerplate.
*/

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
    string[] airLoopNames = new string[] { "Air Loop" };

    string coolerBoilerPlate = @"
EvaporativeCooler:Direct:ResearchSpecial,
  {0},                          !- Name
  {0} Availability,             !- Availability Schedule Name
  0.7 ,                         !- Cooler Design Effectiveness
  ,                             !- Effectiveness Flow Ratio Modifier Curve Name
  autosize,                     !- Primary Air Design Flow Rate m3/s
  30.0 ,                        !- Recirculating Water Pump Design Power
  ,                             !- Water Pump Power Sizing Factor
  ,                             !- Water Pump Power Modifier Curve Name
  {1},                          !- Air Inlet Node Name
  {2},                          !- Air Outlet Node Name
  {2},                          !- Sensor Node Name
  ,                             !- Water Supply Storage Tank Name
  0.0,                          !- Drift Loss Fraction
  3;                            !- Blowdown Concentration Ratio

Schedule:Compact,
   {0} Availability,            ! Name
   Any Number,                  ! Type
   Through: 12/31,              ! Type
   For: AllDays,                ! All days in year
   Until: 24:00,                ! All hours in day
   1;

Schedule:Compact,
   {0} Setpoint,                ! Name
   Any Number,                  ! Type
   Through: 12/31,              ! Type
   For: AllDays,                ! All days in year
   Until: 24:00,                ! All hours in day
   18;

SetpointManager:Scheduled,
  {0} Setpoint Manager,         !- Name
  Temperature,                  !- Control Variable
  {0} Setpoint,                 !- Schedule Name
  {2};                          !- Setpoint Node or NodeList Name

Output:Variable, {0}, Evaporative Cooler Water Volume, Hourly;
Output:Variable, {0}, Evaporative Cooler Electricity Rate, Hourly;
Output:Variable, {0}, Evaporative Cooler Wet Bulb Effectiveness, Hourly;

Output:Variable, {2}, System Node Temperature, Hourly;";


    private IdfObject FindObject(IdfReader idfReader, string objectType, string objectName)
    {
      return idfReader[objectType].First(o => o[0] == objectName);
    }

    public override void BeforeEnergySimulation()
    {
      IdfReader idfReader = new IdfReader(
        ApiEnvironment.EnergyPlusInputIdfPath,
        ApiEnvironment.EnergyPlusInputIddPath
        );


      foreach (string airLoopName in airLoopNames)
      {
        string branchName = airLoopName + " AHU Main Branch";
        IdfObject branch = FindObject(idfReader, "Branch", branchName);

        string coolerName = airLoopName + " AHU Cooler";
        string coolerInletNode = branch[4].Value;
        string coolerOutletNode = airLoopName + " AHU Cooler Outlet Node";

        branch[4].Value = coolerOutletNode;

        string[] newFields = new string[] { "EvaporativeCooler:Direct:ResearchSpecial", coolerName, coolerInletNode, coolerOutletNode };
        branch.InsertFields(2, newFields);

        string cooler = String.Format(coolerBoilerPlate, coolerName, coolerInletNode, coolerOutletNode);

        idfReader.Load(cooler);
      }
      idfReader.Save();
    }
  }
}