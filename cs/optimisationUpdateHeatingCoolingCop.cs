/*
Reads optimisation variables from the "OptimisationVariables" table

Purpose:
This DesignBuilder C# script reads optimisation variables from the DesignBuilder site table named "OptimisationVariables"
and applies them to the EnergyPlus input IDF just before simulation starts.

Main Steps:
1) Read "heatingCOP" and "coolingEER" records from the "OptimisationVariables" table (field: "VariableCurrentValue").
2) Open the generated EnergyPlus IDF using EpNet.
3) If a variable value is "UNKNOWN", a MessageBox is shown and that value is not applied.
4) If a value is not "UNKNOWN", update:
   - Coil:WaterHeating:AirToWaterHeatPump:Pumped (Name = "HP Water Heater HP Water Heating Coil", Field = "Rated COP")
   - Chiller:Electric:EIR (all instances) (Field = "Reference COP")
5) Save the IDF after modifications.

How to Use:

Configuration
- Required record keys:
  - "heatingCOP": used to set "Rated COP" on the named water-heating HP coil
  - "coolingEER": used to set "Reference COP" on Chiller:Electric:EIR objects
- Default object name targeted for heating COP update:
  - "HP Water Heater HP Water Heating Coil" 

Prerequisites / Placeholders
- DesignBuilder Site table "OptimisationVariables" expects keys "heatingCOP" and "coolingEER".
- The EnergyPlus model must contain:
  - Coil:WaterHeating:AirToWaterHeatPump:Pumped object with Name exactly:"HP Water Heater HP Water Heating Coil" (default key)
  - One or more Chiller:Electric:EIR objects (all will be updated)

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. 
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System.Collections.Generic;
using System.Windows.Forms;
using DB.Api;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        public override void BeforeEnergySimulation()
        {
            // Access the current DesignBuilder site and read optimisation variables
            Site site = ApiEnvironment.Site;
            Table optimisationVariablesTable = site.GetTable("OptimisationVariables");
            Record heatingCopRecord = optimisationVariablesTable.Records["heatingCOP"];
            string heatingCopValue = heatingCopRecord["VariableCurrentValue"];

            IdfReader idf = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // Update the specific water heating heat pump coil COP, if present and not UNKNOWN
            IEnumerable<IdfObject> waterHeatingHeatPumpCoils = idf["Coil:WaterHeating:AirToWaterHeatPump:Pumped"];
            foreach (IdfObject coil in waterHeatingHeatPumpCoils)
            {
                // USER CONFIGURATION: Define name of water heater heating coil
                // Only apply to the expected placeholder object by exact Name match
                if (coil["Name"].Equals("HP Water Heater HP Water Heating Coil"))
                {
                    if (heatingCopValue.Equals("UNKNOWN"))
                    {
                        // Skip applying this value if optimisation table contains UNKNOWN
                        MessageBox.Show("Cannot set heating COP: UNKNOWN value in OptimisationVariables table.");
                    }
                    else
                    {
                        // Apply optimisation variable value to the EnergyPlus input field
                        coil["Rated COP"].Value = heatingCopValue;
                    }
                }
            }

            // Read cooling EER value (record key: "coolingEER", field: "VariableCurrentValue")
            Record coolingEerRecord = optimisationVariablesTable.Records["coolingEER"];
            string coolingEerValue = coolingEerRecord["VariableCurrentValue"];

            // Update all Chiller:Electric:EIR objects (Reference COP field), if value is not UNKNOWN
            IEnumerable<IdfObject> electricEirChillers = idf["Chiller:Electric:EIR"];
            foreach (IdfObject chiller in electricEirChillers)
            {
                if (coolingEerValue.Equals("UNKNOWN"))
                {
                    // Skip applying this value if optimisation table contains UNKNOWN
                    MessageBox.Show("Cannot set cooling value: UNKNOWN value in OptimisationVariables table.");
                }
                else
                {
                    // Apply optimisation variable value to the EnergyPlus input field
                    chiller["Reference COP"].Value = coolingEerValue;
                }
            }

            idf.Save();
        }
    }
}