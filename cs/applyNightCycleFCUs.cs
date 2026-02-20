/*
This script adds the "Night Cycle Operation" control for zone fan-coil units.

DesignBuilder automatically applies the unit availability schedule to the child fan.
This setup would not work well for night cycle as the manager overrides only the fan availability.

To allow the night cycling, the script forces the unit schedule to be "ON 24/7".
*/

using System;
using System.Runtime;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        string nightCycleObjects = @"AvailabilityManagerAssignmentList,
    {0},                                  !- Name
    AvailabilityManager:NightCycle,       !- Availability Manager 1 Object Type
    {1} Night Cycle Operation;            !- Availability Manager 1 Name

    AvailabilityManager:NightCycle,
    {1} Night Cycle Operation,            !- Name
    On 24/7,                              !- Applicability Schedule Name
    {3},                                  !- Fan Schedule Name
    CycleOnControlZone,                   !- Control Type
    1,                                    !- Thermostat Tolerance deltaC
    FixedRunTime,                         !- Cycling Run Time Control Type
    3600,                                 !- Cycling Run Time s
    {2};                                  !- Control zone name";

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

        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            StringBuilder idfContent = new StringBuilder();

            IEnumerable<IdfObject> fanCoils = idfReader["ZoneHVAC:FourPipeFanCoil"];
            foreach (IdfObject fanCoil in fanCoils)
            {
                // Add assignment list and nigt cycle manager to the idf
                string fanCoilName = fanCoil["Name"].Value;
                string name = fanCoilName + " Assignment List";
                string zoneName = fanCoilName.Replace(" Fan Coil Unit", "");
                fanCoil.AddField(name, "! - Availability Manager List Name");
                fanCoil["Availability Schedule Name"].Value = "On 24/7";

                // Get fan availability
                IdfObject fan = FindObject(idfReader, fanCoil["Supply Air Fan Object Type"].Value, fanCoil["Supply Air Fan Name"].Value);
                string fanSchedule = fan["Availability Schedule Name"];

                idfContent.AppendFormat(nightCycleObjects, name, fanCoilName, zoneName, fanSchedule);
                idfContent.Append(Environment.NewLine);
            }
            idfReader.Load(idfContent.ToString());
            idfReader.Save();
        }
    }
}
