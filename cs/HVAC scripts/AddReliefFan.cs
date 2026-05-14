/*
Adds a variable-volume relief fan to the outdoor air section of specified air handling units.
Version:1.0

Purpose:
- For each named AHU, this script inserts a Fan:VariableVolume object into the outdoor air
  relief stream, between the OA mixer/controller and the final exhaust node.
- It rewires the OutdoorAir:Mixer relief stream outlet node and Controller:OutdoorAir relief
  outlet node to point to a new intermediate inlet node, while preserving the original terminal
  exhaust node name so that downstream references remain valid.
- The fan is registered in the AirLoopHVAC:OutdoorAirSystem:EquipmentList as the second
  component (after the mixer), so EnergyPlus processes them in the correct simulation order.

Main Steps:
1) Read the existing OutdoorAir:Mixer, Controller:OutdoorAir, and
   AirLoopHVAC:OutdoorAirSystem:EquipmentList objects for each target AHU.
2) Insert a new intermediate inlet node and rewire both the mixer relief stream outlet and the
   OA controller relief outlet to point to it; the original terminal exhaust node is preserved as the fan outlet.
3) Append the new Fan:VariableVolume to the OA equipment list, remove any pre-existing duplicate
   fan objects, load the new fan boilerplate into the IDF, and save.

How to Use:

Configuration
- ahuNames (USER CONFIGURATION SECTION):
    Array of AHU object name strings, e.g. { "Air Loop AHU", "Air Loop 1 AHU" }.
    Each entry must exactly match the Name field of the corresponding IDF objects.
    The air loop base name is derived by stripping the trailing " AHU" suffix.
- fanBoilerplate (USER CONFIGURATION SECTION):
    IDF template for Fan:VariableVolume. Edit fixed performance values here
    (efficiency, pressure rise, flow rate, motor efficiency, power coefficients, etc.)
    as needed for your project. Placeholders {0}-{3} are filled automatically at runtime.

Prerequisites / Placeholders
- The base model must contain the respective AHU objects.

Notes:
- The original terminal exhaust node ("<airLoopName> AHU Relief Air Outlet") is preserved as the
  fan outlet, so any downstream references (e.g. zone exhaust connections) remain valid.
- A new intermediate node ("<airLoopName> Relief Fans Inlet") is created only in memory;
  no explicit Node object is needed in EnergyPlus - node names are resolved by string matching.
- Any existing Fan:VariableVolume with the same generated name is removed before the new one is
  loaded, preventing duplicate object errors on repeated simulation runs.
- Fan performance coefficients in the boilerplate are example values; adjust them to match your
  project's fan curve data.
- This script does NOT modify AirLoopHVAC or any supply-side objects.

DISCLAIMER: This script is provided as-is without warranty. 
DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script.
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Linq;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class AddReliefFan : ScriptBase, IScript
    {
        // ---------------------------
        // USER CONFIGURATION SECTION
        // ---------------------------
        // List of AHU names to process. Must match the IDF object Name fields exactly.
        private readonly string[] ahuNames = new string[]
        {
            "Air Loop AHU",
            "Air Loop 1 AHU"
        };

        // ---------------------------
        // USER CONFIGURATION SECTION
        // ---------------------------
        // IDF boilerplate template for Fan:VariableVolume.
        // Placeholders filled at runtime:
        //   {0} = fan name             (<airLoopName> Relief Fans)
        //   {1} = air inlet node name  (<airLoopName> Relief Fans Inlet - new intermediate node)
        //   {2} = air outlet node name (<airLoopName> AHU Relief Air Outlet - original terminal node)
        //   {3} = air loop base name   (used in inline IDF comments only)
        private readonly string fanBoilerplate = @"Fan:VariableVolume,
  {0},                              !- Name
  On 24/7,                          !- Availability Schedule Name
  0.42,                             !- Fan Total Efficiency
  93.375,                           !- Pressure Rise {{Pa}}
  65.13,                            !- Maximum Flow Rate {{m3/s}}
  Fraction,                         !- Fan Power Minimum Flow Rate Input Method
  0.05,                             !- Fan Power Minimum Flow Fraction
  ,                                 !- Fan Power Minimum Air Flow Rate {{m3/s}}
  0.9,                              !- Motor Efficiency
  1,                                !- Motor In Airstream Fraction
  0.0015302446,                     !- Fan Power Coefficient 1
  0.0052080574,                     !- Fan Power Coefficient 2
  1.1086242000,                     !- Fan Power Coefficient 3
  -0.1163556300,                    !- Fan Power Coefficient 4
  0,                                !- Fan Power Coefficient 5
  {1},                              !- Air Inlet Node Name  ({3} Relief Fans Inlet)
  {2},                              !- Air Outlet Node Name ({3} AHU Relief Air Outlet - original terminal)
  Relief Fans;                      !- End-Use Subcategory";

        // Hook Point: Executes before the EnergyPlus simulation starts
        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            foreach (string ahuName in ahuNames)
            {
                // Derive the air loop base name (strip the " AHU" suffix)
                string airLoopName = ahuName.Replace(" AHU", "");

                // Derived object and node names
                string oaEquipListName = ahuName + " Outdoor air Equipment List";
                string fanName = airLoopName + " Relief Fans";
                string fanOutletNode = airLoopName + " AHU Relief Air Outlet"; // preserved: was the original terminal node
                string fanInletNode = airLoopName + " Relief Fans Inlet";     // new intermediate node inserted before the fan

                // Rewire OutdoorAir:Mixer relief stream outlet - new fan inlet node.
                // Field layout: [0] Name, [1] Mixed Air, [2] OA Inlet, [3] Relief Air Stream, [4] Return Air Stream
                foreach (IdfObject oaMixer in idfReader["OutdoorAir:Mixer"])
                {
                    if (oaMixer["Name"].Value.ToString()
                            .Equals(ahuName + " Outdoor Air Mixer", StringComparison.OrdinalIgnoreCase))
                    {
                        oaMixer[3].Value = fanInletNode;
                    }
                }

                // Rewire Controller:OutdoorAir relief outlet - new fan inlet node.
                // Must stay in sync with the mixer relief node; a mismatch causes a severe EnergyPlus error.
                // Field layout: [0] Name, [1] Relief Air Outlet Node, [2] Return Air, ...
                foreach (IdfObject oaController in idfReader["Controller:OutdoorAir"])
                {
                    if (oaController["Name"].Value.ToString()
                            .Equals(ahuName + " Outdoor Air Controller", StringComparison.OrdinalIgnoreCase))
                    {
                        oaController[1].Value = fanInletNode;
                    }
                }

                // Register the fan as Component 2 in the OA equipment list
                // (Component 1 is the mixer, which populates fanInletNode during simulation)
                foreach (IdfObject oaEquipList in idfReader["AirLoopHVAC:OutdoorAirSystem:EquipmentList"])
                {
                    if (oaEquipList["Name"].Value.ToString()
                            .Equals(oaEquipListName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (oaEquipList.Count > 4)
                        {
                            oaEquipList[3].Value = "Fan:VariableVolume";
                            oaEquipList[4].Value = fanName;
                        }
                        else
                        {
                            oaEquipList.AddFields(new string[] { "Fan:VariableVolume", fanName });
                        }
                    }
                }

                // Remove any pre-existing fan with the same name to avoid duplicate object errors
                foreach (IdfObject existingFan in idfReader["Fan:VariableVolume"]
                    .Where(f => f["Name"].Value.ToString()
                        .Equals(fanName, StringComparison.OrdinalIgnoreCase))
                    .ToList())
                {
                    idfReader.Remove(existingFan);
                }

                // Load the new Fan:VariableVolume into the IDF
                idfReader.Load(String.Format(fanBoilerplate,
                    fanName,       // {0} fan name
                    fanInletNode,  // {1} air inlet node (new intermediate node)
                    fanOutletNode, // {2} air outlet node (original terminal, preserved)
                    airLoopName)); // {3} air loop base name (IDF comments only)
            }

            idfReader.Save();
        }
    }
}