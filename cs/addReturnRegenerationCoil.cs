/*
Add Coil:Heating:Electric to the air loop return air stream.

The script adds the component to air loops specified in "airLoopNames" array.
Coil parameters and setpoint can be set in the object boilerplate.
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

        string coilBoilerPlate = @"Coil:Heating:Electric,
  {0},                          !- Name
  {0} Availability,             !- Availability Schedule Name
  1.00,                         !- Efficiency (%)
  autosize,                     !- Nominal capacity (W)
  {1},                          !- Air Inlet Node Name
  {2},                          !- Air Outlet Node Name
  {2};                          !- Sensor Node Name

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

Output:Variable, {0}, Heating Coil Electricity Rate, Hourly;
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

                string coilName = airLoopName + " AHU Regeneration Coil";
                string coilInletNode = branch[4].Value;
                string coilOutletNode = airLoopName + " AHU Regeneration Coil Outlet Node";

                branch[4].Value = coilOutletNode;

                string[] newFields = new string[] { "Coil:Heating:Electric", coilName, coilInletNode, coilOutletNode };
                branch.InsertFields(2, newFields);

                string coil = String.Format(coilBoilerPlate, coilName, coilInletNode, coilOutletNode);

                idfReader.Load(coil);
            }
            idfReader.Save();
        }
    }
}