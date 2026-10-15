/*
Replace the supply fan of tagged AHUs with a two-speed Fan:SystemModel.

Purpose:
- Replaces the supply fan of each tagged AHU with a Fan:SystemModel with two discrete speeds, as in the PNNL
  ASHRAE 90.1-2019 prototype models (Speed 1 = 66 % flow / 40 % power, Speed 2 = 100 % flow / 100 % power). This
  represents the two-stage fan control of ASHRAE 90.1-2019 section 6.5.3.2.1(a).
- Applies to the Unitary System and the General AHU (including DOAS).
- Unitary heat pump, furnace and unitary heat-cool AHUs are not supported, because EnergyPlus does not accept a
  Fan:SystemModel in these objects.

Main Steps:
1) Find the air loops whose name ends with the tag (default "FSM"). DesignBuilder names the AHU after its air loop,
   e.g. "Air Loop FSM Unitary System", so the air loop name identifies the AHU.
2) Unitary System: replace its supply fan and set its no-load (ventilation-only) supply air flow to the low-speed flow.
3) General AHU: replace the fan on the main branch downstream of the outdoor air system (the supply fan). Fans upstream
   of the outdoor air system are return or extract fans and are not changed.
4) The Fan:SystemModel keeps the name, nodes, availability schedule, maximum flow, pressure rise, efficiencies and end-use
   subcategory of the replaced fan. Save the IDF and show a summary.

How to Use:

Configuration
- airLoopNameTag: text at the end of the names of the air loops whose AHU is converted (not case-sensitive).
- fanSystemModelBoilerplate: the Fan:SystemModel written for each converted fan. Edit the speed fractions there if needed.
- adjustNoLoadFlow / noLoadFlowFraction: Unitary System only. The no-load supply air flow is set to this fraction of the
  fan flow so that the low speed is used. Keep noLoadFlowFraction equal to the Speed 1 Flow Fraction.
- showMessages: set to false for parametric, optimisation or batch runs to avoid a dialog on every simulation.

Prerequisites / Placeholders
- Detailed HVAC. Add FSM at the end of the air loop name of each AHU to convert, e.g. "Air Loop FSM".
- Any DesignBuilder supply fan type can be used as the placeholder; its settings are carried over.
- Unitary System: set the fan operating mode schedule to continuous during occupied hours, otherwise the fan is off when
  there is no load and the low speed is not used.

Notes:
- Tag only AHUs that should have a two-speed fan. On a variable volume AHU, the two speeds replace the variable-speed
  part-load curve.
- General AHU with constant volume terminals: the air loop flow does not drop, so the fan stays at Speed 2.
- Unitary System with single-speed DX coils: the compressor runs at full airflow, so the low speed is only used while the
  compressor is off.
- Night ventilation fan performance (FanPerformance:NightVentilation) and AirflowNetwork fan references to a converted fan
  are not updated.
- Field positions match EnergyPlus 25.1 (DesignBuilder 2026.1).

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script.
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class ReplaceAhuFanWithFanSystemModel : ScriptBase, IScript
    {
        // ---------------------------
        // USER CONFIGURATION SECTION
        // ---------------------------
        // The AHU of each air loop whose name ends with this text is converted (not case-sensitive).
        private readonly string airLoopNameTag = "FSM";

        // Unitary System only: set the no-load (ventilation-only) supply air flow to the low-speed flow.
        // Keep noLoadFlowFraction equal to the Speed 1 Flow Fraction in the boilerplate below.
        private readonly bool adjustNoLoadFlow = true;
        private readonly double noLoadFlowFraction = 0.66;

        // Show a summary after the IDF is updated (set to false for parametric, optimisation or batch runs).
        private readonly bool showMessages = true;

        // Fan:SystemModel written for each converted fan: two discrete speeds as in the PNNL ASHRAE 90.1-2019 prototype models.
        // {0}-{9} are copied from the replaced fan. The speed power fractions define part-load power, so no curve is needed.
        private readonly string fanSystemModelBoilerplate = @"
Fan:SystemModel,
    {0},                          !- Name
    {1},                          !- Availability Schedule Name
    {2},                          !- Air Inlet Node Name
    {3},                          !- Air Outlet Node Name
    {4},                          !- Design Maximum Air Flow Rate {{m3/s}}
    Discrete,                     !- Speed Control Method
    0.0,                          !- Electric Power Minimum Flow Rate Fraction
    {5},                          !- Design Pressure Rise {{Pa}}
    {6},                          !- Motor Efficiency
    {7},                          !- Motor In Air Stream Fraction
    autosize,                     !- Design Electric Power Consumption {{W}}
    TotalEfficiencyAndPressure,   !- Design Power Sizing Method
    ,                             !- Electric Power Per Unit Flow Rate {{W/(m3/s)}}
    ,                             !- Electric Power Per Unit Flow Rate Per Unit Pressure {{W/((m3/s)-Pa)}}
    {8},                          !- Fan Total Efficiency
    ,                             !- Electric Power Function of Flow Fraction Curve Name
    ,                             !- Night Ventilation Mode Pressure Rise {{Pa}}
    ,                             !- Night Ventilation Mode Flow Fraction
    ,                             !- Motor Loss Zone Name
    ,                             !- Motor Loss Radiative Fraction
    {9},                          !- End-Use Subcategory
    2,                            !- Number of Speeds
    0.66,                         !- Speed 1 Flow Fraction
    0.40,                         !- Speed 1 Electric Power Fraction
    1.0,                          !- Speed 2 Flow Fraction
    1.0;                          !- Speed 2 Electric Power Fraction
";
        // ---------------------------
        // END OF USER CONFIGURATION SECTION
        // ---------------------------

        // Positions (0 = Name) of the fields copied from each DesignBuilder fan type, in boilerplate order {1}-{9}:
        // Availability Schedule, Air Inlet Node, Air Outlet Node, Maximum Flow Rate, Pressure Rise,
        // Motor Efficiency, Motor In Airstream Fraction, Fan Total Efficiency, End-Use Subcategory.
        private readonly Dictionary<string, int[]> fanFields = new Dictionary<string, int[]>(StringComparer.OrdinalIgnoreCase)
        {
            { "Fan:ConstantVolume", new[] { 1, 7, 8, 4, 3, 5, 6, 2, 9 } },
            { "Fan:OnOff", new[] { 1, 7, 8, 4, 3, 5, 6, 2, 11 } },
            { "Fan:VariableVolume", new[] { 1, 15, 16, 4, 3, 8, 9, 2, 17 } }
        };

        private IdfReader idf;
        private List<string> summary = new List<string>();

        public override void BeforeEnergySimulation()
        {
            idf = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);
            summary.Clear();
            List<string> airLoopNames = new List<string>();

            // AirLoopHVAC: [4] Branch List Name. BranchList: branch names from [1].
            foreach (IdfObject airLoop in idf["AirLoopHVAC"].ToList())
            {
                string airLoopName = airLoop[0].Value.Trim();
                airLoopNames.Add(airLoopName);
                if (!airLoopName.EndsWith(airLoopNameTag, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                IdfObject branchList = FindObject("BranchList", airLoop[4].Value);
                for (int i = 1; i < branchList.Count; i++)
                {
                    ConvertAhu(FindObject("Branch", branchList[i].Value), airLoopName);
                }
            }

            if (summary.Count > 0)
            {
                idf.Save();
            }
            else
            {
                summary.Add("No air loop name ends with '" + airLoopNameTag + "', nothing was changed. Air loops found: " + string.Join(", ", airLoopNames));
            }

            if (showMessages)
            {
                MessageBox.Show(string.Join("\n", summary), "Replace AHU fan with Fan:SystemModel");
            }
        }

        // Branch: [0] Name, [1] Pressure Drop Curve, then (Object Type, Name, Inlet Node, Outlet Node) groups from [2].
        private void ConvertAhu(IdfObject branch, string ahuName)
        {
            bool afterOutdoorAirSystem = false;
            for (int i = 2; i + 1 < branch.Count; i += 4)
            {
                string componentType = branch[i].Value.Trim();
                string componentName = branch[i + 1].Value.Trim();

                if (componentType.Equals("AirLoopHVAC:OutdoorAirSystem", StringComparison.OrdinalIgnoreCase))
                {
                    afterOutdoorAirSystem = true;
                }
                else if (componentType.Equals("AirLoopHVAC:UnitarySystem", StringComparison.OrdinalIgnoreCase))
                {
                    ConvertUnitarySystem(FindObject(componentType, componentName), ahuName);
                }
                else if (afterOutdoorAirSystem && fanFields.ContainsKey(componentType))
                {
                    // General AHU supply fan: replace the fan and its type in the branch.
                    ReplaceFan(componentType, componentName, ahuName);
                    branch[i].Value = "Fan:SystemModel";
                }
            }
        }

        // AirLoopHVAC:UnitarySystem: [7] Supply Fan Object Type, [8] Supply Fan Name,
        // [31] No Load Supply Air Flow Rate Method, [32] No Load Supply Air Flow Rate,
        // [34] No Load Fraction of Autosized Cooling Supply Air Flow Rate.
        private void ConvertUnitarySystem(IdfObject unitary, string ahuName)
        {
            string fanType = unitary[7].Value.Trim();
            if (!fanFields.ContainsKey(fanType))
            {
                return; // no supply fan, or already a Fan:SystemModel
            }

            string fanFlow = ReplaceFan(fanType, unitary[8].Value.Trim(), ahuName);
            unitary[7].Value = "Fan:SystemModel";

            if (!adjustNoLoadFlow)
            {
                return;
            }

            // Run the fan at low speed when there is no heating or cooling load (ventilation only).
            double flow;
            if (double.TryParse(fanFlow, NumberStyles.Float, CultureInfo.InvariantCulture, out flow))
            {
                // Hard-sized fan: fixed no-load flow.
                unitary[31].Value = "SupplyAirFlowRate";
                unitary[32].Value = (flow * noLoadFlowFraction).ToString("0.######", CultureInfo.InvariantCulture);
            }
            else
            {
                // Autosized fan: no-load flow as a fraction of the autosized cooling flow.
                unitary[31].Value = "FractionOfAutosizedCoolingValue";
                unitary[32].Value = "";
                unitary[34].Value = noLoadFlowFraction.ToString(CultureInfo.InvariantCulture);
            }
        }

        // Replaces a fan with a Fan:SystemModel of the same name and returns the fan maximum flow rate.
        private string ReplaceFan(string fanType, string fanName, string ahuName)
        {
            IdfObject oldFan = FindObject(fanType, fanName);
            int[] positions = fanFields[fanType];

            string[] values = new string[positions.Length + 1];
            values[0] = oldFan[0].Value.Trim();
            for (int k = 0; k < positions.Length; k++)
            {
                values[k + 1] = positions[k] < oldFan.Count ? oldFan[positions[k]].Value.Trim() : "";
            }

            idf.Remove(oldFan);
            idf.Load(string.Format(fanSystemModelBoilerplate, values));
            summary.Add(ahuName + ": " + fanName + " (" + fanType + ") replaced with Fan:SystemModel.");
            return values[4];
        }

        private IdfObject FindObject(string objectType, string objectName)
        {
            return idf[objectType].First(o =>
                o[0].Value.Trim().Equals(objectName.Trim(), StringComparison.OrdinalIgnoreCase));
        }
    }
}