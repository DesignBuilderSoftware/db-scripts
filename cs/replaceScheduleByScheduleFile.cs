/*
This script replaces Schedule:Compact with schedule:File objects in the IDF file.

*/
using System.Runtime;
using System.Linq;
using System.Collections.Generic;
using System;
using System.Windows.Forms;
using DB.Extensibility.Contracts;
using EpNet;


namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        private IdfReader reader;

        public override void BeforeEnergySimulation()
        {
            reader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // define schedules to be replaced
            ReplaceCompactSchedule(
                scheduleName: "test schedule",
                filePath: @"C:\path\to\your\file.csv",
                columnNumber: 1,
                rowsToSkip: 0);

            reader.Save();
        }
        public IdfObject FindObject(string objectType, string objectName)
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

        public void ReplaceCompactSchedule(string scheduleName, string filePath, int columnNumber, int rowsToSkip, string columnSeparator = "comma")
        {
            IdfObject schedule = FindObject("Schedule:Compact", scheduleName);
            string scheduleFile = GetScheduleFile(scheduleName, filePath, columnNumber, rowsToSkip, columnSeparator);

            reader.Load(scheduleFile);
            reader.Remove(schedule);
        }


        public string GetScheduleFile(string scheduleName, string filePath, int columnNumber, int rowsToSkip, string columnSeparator)
        {
            string scheduleTemplate = @"Schedule:File,
  {0},                !- Name
  Any Number,         !- Schedule Type Limits Name
  {1},                !- File Name
  {2},                !- Column Number
  {3},                !- Rows to Skip at Top
  8760,               !- Number of Hours of Data
  {4},                !- Column Separator
  ,                   !- Interpolate to Timestep
  60;                 !- Minutes per Item";
            return String.Format(scheduleTemplate, scheduleName, filePath, columnNumber, rowsToSkip, columnSeparator);
        }
    }
}
