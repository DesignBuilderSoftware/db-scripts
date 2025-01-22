/*
This script applies the "AvailabilityManager:NightVentilation" control to specified air loops.

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
        // Define IDF templates
        readonly string nightVentilationTemplate = @"AvailabilityManager:NightVentilation,
    {0},                     !- Name
    Night Ventilation Applicability Schedule, !- Applicability Schedule Name
    {1},                     !- Fan Schedule Name
    Night Ventilation Setpoint Schedule,  !- Ventilation Temperature Schedule Name
    2,                       !- Ventilation Temperature Difference deltaC
    15,                      !- Ventilation Temperature Low Limit C
    0.5,                     !- Night Venting Flow Fraction
    {2};                     !- Control Zone Name";

        readonly string commonObjects = @"Schedule:Compact,
    Night Ventilation Setpoint Schedule,  !- Name
    Any Number,              !- Schedule Type Limits Name
    Through: 12/31,          !- Field 1
    For: AllDays,            !- Field 2
    Until: 24:00,            !- Field 3
    24;                      !- Field 4

Schedule:Compact,
    Night Ventilation Applicability Schedule,  !- Name
    Any Number,              !- Schedule Type Limits Name
    Through: 12/31,          !- Field 1
    For: AllDays,            !- Field 2
    Until: 24:00,            !- Field 3
    1;                       !- Field 4";

        private IdfObject FindObject(IdfReader reader, string objectType, string objectName)
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

        private string FindFanScheduleName(IdfReader reader, string airLoopName)
        {
            IdfObject mainBranch = FindObject(reader, "Branch", airLoopName + " AHU Main Branch");
            string[] fanTypes = { "Fan:VariableVolume", "Fan:ConstantVolume", "Fan:OnOff" };
            var result = mainBranch
                .Select((field, idx) => new { field, idx })
                .FirstOrDefault(x => fanTypes.Contains(x.field.Value));
            if (result == null)
            {
                throw new Exception("Cannot find fan in air loop: " + airLoopName);
            }

            int fanTypeIndex = result.idx;
            string fanType = mainBranch[fanTypeIndex].Value;
            string fanName = mainBranch[fanTypeIndex + 1].Value;
            IdfObject fan = FindObject(reader, fanType, fanName);
            return fan[1].Value;
        }

        private void UpdateAssignmentList(IdfReader reader, string airLoopName, string managerCls, string managerName)
        {
            string availabilityAssignmentListCls = "AvailabilityManagerAssignmentList";
            string availabilityAssignmentListName = airLoopName + " AvailabilityManager List";

            IdfObject availabilityAssignmentList = FindObject(reader, availabilityAssignmentListCls, availabilityAssignmentListName);
            if (availabilityAssignmentList[1] == "AvailabilityManager:Scheduled")
            {
                availabilityAssignmentList[1].Value = managerCls;
                availabilityAssignmentList[2].Value = managerName;
            }
            else
            {
                availabilityAssignmentList.AddFields(managerCls, managerName);
            }
        }

        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // Define key (air loop name) value (control zone name) pairs
            // Check IDF to get exact names
            var loopControlZonePairs = new Dictionary<string, string>
             {
                 { "Air Loop", "Block1:Zone1" },
                 // { "Air Loop 1", "Block1:Zone2" }
             };

            StringBuilder idfContent = new StringBuilder();

            foreach (var pair in loopControlZonePairs)
            {    
                string airLoopName = pair.Key;
                string controlZoneName = pair.Value;

                string nightVentilationCls = "AvailabilityManager:NightVentilation";
                string nightVentilationName = airLoopName + " Night Ventilation";

                string fanScheduleName = FindFanScheduleName(idfReader, airLoopName);
                idfContent.AppendFormat(nightVentilationTemplate, nightVentilationName, fanScheduleName, controlZoneName);
                idfContent.Append(Environment.NewLine);

                UpdateAssignmentList(idfReader, airLoopName, nightVentilationCls, nightVentilationName);
            }

            idfContent.Append(commonObjects);

            idfReader.Load(idfContent.ToString());
            idfReader.Save();
        }
    }
}