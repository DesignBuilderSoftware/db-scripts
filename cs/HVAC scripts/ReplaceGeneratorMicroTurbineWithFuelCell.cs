/*
Replace Generator:MicroTurbine with a Generator:FuelCell (CHP) on a hot-water loop.

Purpose:
- This DesignBuilder C# script swaps a Generator:MicroTurbine placeholder for a full
  Generator:FuelCell cogeneration unit. The MicroTurbine is used only as an anchor: the
  script reads its plant connection (the hot-water heat-recovery nodes it sits on) and its
  electrical registration, then injects the complete Generator:FuelCell object family and
  rewires the loop and the electric load centre to point at it. The result is a fuel cell
  that recovers heat into the same HW loop branch the MicroTurbine occupied and that reports
  its electricity through the same ElectricLoadCenter.

Main Steps:
1) For each generator name in the configured list, find that Generator:MicroTurbine object and
   the Branch it sits on, reading the branch component inlet/outlet node names (these are the
   HW heat-recovery water nodes).
2) Build the Generator:FuelCell family boilerplate (power module, air/fuel/water supply,
   auxiliary heater, exhaust-gas-to-water heat exchanger, electrical storage, inverter, and
   their performance curves), inheriting the discovered water nodes on the heat exchanger.
3) Rewire the plant side (Branch + PlantEquipmentList) so the component is now the fuel cell
   exhaust heat exchanger, and the electric side (ElectricLoadCenter:Generators) so the
   generator object type is now Generator:FuelCell; load the boilerplate and remove the
   MicroTurbine placeholder; save the IDF.

How to Use:

Configuration
- USER CONFIGURATION SECTION (in BeforeEnergySimulation): generatorNamesToReplace - the list of
  Generator:MicroTurbine names to swap (one entry per generator). Only the generators named here
  are replaced; any other Generator:MicroTurbine objects in the model are left untouched.
- USER CONFIGURATION SECTION (in BeforeEnergySimulation): skinLossZoneNames - an array of zones,
  index-aligned with generatorNamesToReplace (same order). For each generator: leave the entry
  empty ("") to send its skin/auxiliary heat losses to ambient (removed from the building, the
  default), or give a zone name to credit those losses to that zone. Entries omitted past the end
  of the array default to ambient.
- USER CONFIGURATION SECTION (in GetFuelCellIdf): all fuel-cell performance values
  (efficiencies, operating points, curves, fuel composition, heat-exchanger coefficients) are
  in one boilerplate string. Edit there to match the specific unit being modelled. Values
  shipped here are taken from the EnergyPlus MicroCogeneration SOFC reference example.

Prerequisites / Placeholders
The base model must contain:
- One or more Generator:MicroTurbine objects, each placed via the DesignBuilder interface on a
  hot-water loop supply branch (heat recovery enabled, i.e. with HW inlet/outlet nodes), in
  parallel with the boiler. The names of the ones to replace go in generatorNamesToReplace.
- An ElectricLoadCenter:Distribution / ElectricLoadCenter:Generators pair already referencing
  those generators (create these in DesignBuilder GUI in the Generator tab at Building level).
- A Zone is only needed if you choose to credit a generator's skin losses to a space via
  skinLossZoneNames; by default the losses go to ambient and no zone is required.

Notes:
- By default, fuel-cell skin and auxiliary heat losses are sent to ambient (removed from the
  building): the PowerModule Zone Name is left blank and the AuxiliaryHeater Skin Loss
  Destination is AirInletForFuelCell. Supplying a zone in skinLossZoneNames instead routes that
  generator's losses into the named zone (PowerModule Zone Name set, AuxiliaryHeater
  Skin Loss Destination SurroundingZone).
- The fuel cell draws combustion and dilution air from newly created OutdoorAir nodes, so it
  does not depend on any zone-exhaust node plumbing.
- The ElectricLoadCenter:Distribution buss type is left untouched (AlternatingCurrent); the
  Generator:FuelCell:Inverter handles DC-to-AC internally.
- Only the Generator:MicroTurbine object is removed. Its old combustion-air OutdoorAir:NodeList
  node is harmless if left orphaned and is intentionally not touched (minimal change).
- The script replaces only the generators named in generatorNamesToReplace; it relies on each
  name matching its Generator:MicroTurbine exactly, and on DesignBuilder's consistent
  "<Generator Name> + ..." naming for the derived sub-objects and nodes.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script.
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Globalization;
using DB.Extensibility.Contracts;
using DB.Api;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class ReplaceMicroTurbineWithFuelCell : ScriptBase, IScript
    {
        private IdfReader idfReader;

        public override void BeforeEnergySimulation()
        {
            idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // ----------------------------
            // USER CONFIGURATION SECTION
            // ----------------------------
            // Placeholder object type to look for and replace.
            string microTurbinePlaceholderType = "Generator:MicroTurbine";

            // Names of the Generator:MicroTurbine objects to replace.
            // Add one entry per generator you want swapped.
            string[] generatorNamesToReplace = new string[]
            {
                "Generator 1",
            };

            // Skin-loss zone for each generator above, in the SAME ORDER (index-aligned with generatorNamesToReplace).
            // Leave an entry empty ("") to send that generator's skin/auxiliary losses to ambient (the default). 
            // Provide a zone name to credit the losses to that zone instead.
            string[] skinLossZoneNames = new string[]
            {
                "",
            };

            ApplyReplacement(microTurbinePlaceholderType, generatorNamesToReplace, skinLossZoneNames);

            idfReader.Save();
        }

        public void ApplyReplacement(
            string microTurbinePlaceholderType,
            string[] generatorNamesToReplace,
            string[] skinLossZoneNames)
        {
            // Apply the swap to each requested generator, pairing it with its skin-loss zone by position.
            // A missing or empty zone entry means "send losses to ambient".
            for (int i = 0; i < generatorNamesToReplace.Length; i++)
            {
                string generatorName = generatorNamesToReplace[i];
                string skinLossZoneName = (i < skinLossZoneNames.Length) ? skinLossZoneNames[i] : "";
                ReplaceGenerator(microTurbinePlaceholderType, generatorName, skinLossZoneName);
            }
        }

        private void ReplaceGenerator(string microTurbinePlaceholderType, string generatorName, string skinLossZoneName)
        {
            // Target object types created/referenced by the swap.
            string fuelCellType = "Generator:FuelCell";
            string exhaustHxType = "Generator:FuelCell:ExhaustGasToWaterHeatExchanger";

            // 1. Locate the named placeholder generator.
            IdfObject microTurbine = FindByName(microTurbinePlaceholderType, generatorName);

            // Derived name for the fuel-cell exhaust heat exchanger (the plant-connected part).
            string hxName = generatorName + " FC Exhaust HX";

            // 2. Locate the branch the placeholder sits on and read its HW heat-recovery nodes.
            string[] waterNodes = GetBranchComponentNodes(microTurbinePlaceholderType, generatorName);
            string waterInletNode = waterNodes[0];
            string waterOutletNode = waterNodes[1];

            // 3. Plant-side rewiring: the loop component becomes the fuel-cell exhaust heat
            //    exchanger (type AND name change). 
            //    Node names on the branch are preserved and are inherited by the heat exchanger below.
            ReplaceObjectTypeInList("Branch", microTurbinePlaceholderType, generatorName, exhaustHxType, hxName);
            ReplaceObjectTypeInList("PlantEquipmentList", microTurbinePlaceholderType, generatorName, exhaustHxType, hxName);

            // 4. Electric-side rewiring: in ElectricLoadCenter:Generators only the object TYPE changes; 
            //    The generator name is preserved. Scoped to this generator name so
            //    other (unlisted) MicroTurbine generators sharing the list are left untouched.
            ReplaceTypeForNamedGenerator("ElectricLoadCenter:Generators", generatorName, microTurbinePlaceholderType, fuelCellType);

            // 5. Inject the Generator:FuelCell family (inheriting the discovered water nodes).
            string fuelCellIdf = GetFuelCellIdf(generatorName, waterInletNode, waterOutletNode, skinLossZoneName);
            idfReader.Load(fuelCellIdf);

            // 6. Remove the placeholder MicroTurbine now that all references point elsewhere.
            idfReader.Remove(microTurbine);
        }

        // ----------------------------
        // USER CONFIGURATION SECTION
        // ----------------------------
        // Full Generator:FuelCell family boilerplate. Tokens replaced at runtime:
        //   __FC__              -> generator / fuel-cell base name (also prefixes all child objects)
        //   __WATER_IN__        -> HW heat-recovery water inlet node (inherited from the placeholder)
        //   __WATER_OUT__       -> HW heat-recovery water outlet node (inherited from the placeholder)
        //   __POWERMODULE_ZONE__-> PowerModule Zone Name (blank = skin loss to ambient)
        //   __AUX_SKIN_DEST__   -> AuxiliaryHeater Skin Loss Destination
        //                          (AirInletForFuelCell when no zone, SurroundingZone when a zone is given)
        //   __AUX_ZONE__        -> AuxiliaryHeater Zone Name to Receive Skin Losses (blank when no zone)
        public string GetFuelCellIdf(
            string generatorName,
            string waterInletNode,
            string waterOutletNode,
            string zoneName)
        {
            string template = @"
Generator:FuelCell,
    __FC__,                          !- Name
    __FC__ FC Power Module,          !- Power Module Name
    __FC__ FC Air Supply,            !- Air Supply Name
    __FC__ FC Fuel Supply,           !- Fuel Supply Name
    __FC__ FC Water Supply,          !- Water Supply Name
    __FC__ FC Auxiliary Heater,      !- Auxiliary Heater Name
    __FC__ FC Exhaust HX,            !- Heat Exchanger Name
    __FC__ FC Battery,               !- Electrical Storage Name
    __FC__ FC Inverter;              !- Inverter Name

Generator:FuelCell:PowerModule,
    __FC__ FC Power Module,          !- Name
    Annex42,                         !- Efficiency Curve Mode
    __FC__ FC Power Curve,           !- Efficiency Curve Name
    0.354,                           !- Nominal Efficiency
    3400,                            !- Nominal Electrical Power {W}
    0,                               !- Number of Stops at Start of Simulation
    0.0,                             !- Cycling Performance Degradation Coefficient
    0,                               !- Number of Run Hours at Beginning of Simulation {hr}
    0.0,                             !- Accumulated Run Time Degradation Coefficient
    10000,                           !- Run Time Degradation Initiation Time Threshold {hr}
    1.4,                             !- Power Up Transient Limit {W/s}
    0.2,                             !- Power Down Transient Limit {W/s}
    0.0,                             !- Start Up Time {s}
    0.2,                             !- Start Up Fuel {kmol}
    ,                                !- Start Up Electricity Consumption {J}
    0.0,                             !- Start Up Electricity Produced {J}
    0.0,                             !- Shut Down Time {s}
    0.2,                             !- Shut Down Fuel {kmol}
    ,                                !- Shut Down Electricity Consumption {J}
    0.0,                             !- Ancillary Electricity Constant Term
    0.0,                             !- Ancillary Electricity Linear Term
    ConstantRate,                    !- Skin Loss Calculation Mode
    __POWERMODULE_ZONE__,            !- Zone Name (blank = skin loss removed to ambient)
    0.6392,                          !- Skin Loss Radiative Fraction (unused when Zone Name blank)
    729,                             !- Constant Skin Loss Rate {W}
    0.0,                             !- Skin Loss U-Factor Times Area Term {W/K}
    ,                                !- Skin Loss Quadratic Curve Name
    6.156E-3,                        !- Dilution Air Flow Rate {kmol/s}
    2307,                            !- Stack Heat loss to Dilution Air {W}
    __FC__ FC Dilution Inlet Node,   !- Dilution Inlet Air Node Name
    __FC__ FC Dilution Outlet Node,  !- Dilution Outlet Air Node Name
    3010,                            !- Minimum Operating Point {W}
    3728;                            !- Maximum Operating Point {W}

Curve:Quadratic,
    __FC__ FC Power Curve,           !- Name
    0.642388,                        !- Coefficient1 Constant
    -1.619E-4,                       !- Coefficient2 x
    2.26007E-8,                      !- Coefficient3 x**2
    0.0,                             !- Minimum Value of x
    10000;                           !- Maximum Value of x

Generator:FuelCell:AirSupply,
    __FC__ FC Air Supply,            !- Name
    __FC__ FC Air Inlet Node,        !- Air Inlet Node Name
    __FC__ FC Blower Power Curve,     !- Blower Power Curve Name
    1.0,                             !- Blower Heat Loss Factor
    QuadraticFunctionofElectricPower, !- Air Supply Rate Calculation Mode
    ,                                !- Stoichiometric Ratio
    __FC__ FC Excess Air Ratio Curve, !- Air Rate Function of Electric Power Curve Name
    2.83507E-3,                      !- Air Rate Air Temperature Coefficient
    ,                                !- Air Rate Function of Fuel Rate Curve Name
    NoRecovery,                      !- Air Intake Heat Recovery Mode
    UserDefinedConstituents,         !- Air Supply Constituent Mode
    5,                               !- Number of UserDefined Constituents
    Nitrogen,                        !- Constituent 1 Name
    0.7728,                          !- Molar Fraction 1
    Oxygen,                          !- Constituent 2 Name
    0.2073,                          !- Molar Fraction 2
    Water,                           !- Constituent 3 Name
    0.0104,                          !- Molar Fraction 3
    Argon,                           !- Constituent 4 Name
    0.0092,                          !- Molar Fraction 4
    CarbonDioxide,                   !- Constituent 5 Name
    0.0003;                          !- Molar Fraction 5

Curve:Cubic,
    __FC__ FC Blower Power Curve,     !- Name
    0.0,                             !- Coefficient1 Constant
    0.0,                             !- Coefficient2 x
    0.0,                             !- Coefficient3 x**2
    0.0,                             !- Coefficient4 x**3
    0.0,                             !- Minimum Value of x
    100000;                          !- Maximum Value of x

Curve:Quadratic,
    __FC__ FC Excess Air Ratio Curve, !- Name
    1.50976E-3,                      !- Coefficient1 Constant
    -7.76656E-7,                     !- Coefficient2 x
    1.30317E-10,                     !- Coefficient3 x**2
    0.0,                             !- Minimum Value of x
    1000000.0;                       !- Maximum Value of x

Generator:FuelSupply,
    __FC__ FC Fuel Supply,           !- Name
    TemperatureFromAirNode,          !- Fuel Temperature Modeling Mode
    __FC__ FC Air Inlet Node,        !- Fuel Temperature Reference Node Name
    ,                                !- Fuel Temperature Schedule Name
    __FC__ Null Cubic,               !- Compressor Power Multiplier Function of Fuel Rate Curve Name
    1.0,                             !- Compressor Heat Loss Factor
    GaseousConstituents,             !- Fuel Type
    ,                                !- Liquid Generic Fuel Lower Heating Value {kJ/kg}
    ,                                !- Liquid Generic Fuel Higher Heating Value {kJ/kg}
    ,                                !- Liquid Generic Fuel Molecular Weight {g/mol}
    ,                                !- Liquid Generic Fuel CO2 Emission Factor
    8,                               !- Number of Constituents in Gaseous Constituent Fuel Supply
    METHANE,                         !- Constituent 1 Name
    0.9490,                          !- Constituent 1 Molar Fraction
    CarbonDioxide,                   !- Constituent 2 Name
    0.0070,                          !- Constituent 2 Molar Fraction
    NITROGEN,                        !- Constituent 3 Name
    0.0160,                          !- Constituent 3 Molar Fraction
    ETHANE,                          !- Constituent 4 Name
    0.0250,                          !- Constituent 4 Molar Fraction
    PROPANE,                         !- Constituent 5 Name
    0.0020,                          !- Constituent 5 Molar Fraction
    BUTANE,                          !- Constituent 6 Name
    0.0006,                          !- Constituent 6 Molar Fraction
    PENTANE,                         !- Constituent 7 Name
    0.0002,                          !- Constituent 7 Molar Fraction
    OXYGEN,                          !- Constituent 8 Name
    0.0002;                          !- Constituent 8 Molar Fraction

Generator:FuelCell:WaterSupply,
    __FC__ FC Water Supply,          !- Name
    __FC__ Null Quadratic,           !- Reformer Water Flow Rate Function of Fuel Rate Curve Name
    __FC__ Null Cubic,               !- Reformer Water Pump Power Function of Fuel Rate Curve Name
    0.0,                             !- Pump Heat Loss Factor
    TemperatureFromAirNode,          !- Water Temperature Modeling Mode
    __FC__ FC Air Inlet Node;        !- Water Temperature Reference Node Name

Curve:Quadratic,
    __FC__ Null Quadratic,           !- Name
    0.0,                             !- Coefficient1 Constant
    0.0,                             !- Coefficient2 x
    0.0,                             !- Coefficient3 x**2
    -1.0E+10,                        !- Minimum Value of x
    1.0E+10;                         !- Maximum Value of x

Curve:Cubic,
    __FC__ Null Cubic,               !- Name
    0.0,                             !- Coefficient1 Constant
    0.0,                             !- Coefficient2 x
    0.0,                             !- Coefficient3 x**2
    0.0,                             !- Coefficient4 x**3
    -1.0E+10,                        !- Minimum Value of x
    1.0E+10;                         !- Maximum Value of x

Generator:FuelCell:AuxiliaryHeater,
    __FC__ FC Auxiliary Heater,      !- Name
    0.0,                             !- Excess Air Ratio
    0.0,                             !- Ancillary Power Constant Term
    0.0,                             !- Ancillary Power Linear Term
    0.5,                             !- Skin Loss U-Factor Times Area Value {W/K}
    __AUX_SKIN_DEST__,               !- Skin Loss Destination
    __AUX_ZONE__,                    !- Zone Name to Receive Skin Losses (blank when to air inlet)
    Watts,                           !- Heating Capacity Units
    0.0,                             !- Maximum Heating Capacity in Watts {W}
    0.0;                             !- Minimum Heating Capacity in Watts {W}

Generator:FuelCell:ExhaustGasToWaterHeatExchanger,
    __FC__ FC Exhaust HX,            !- Name
    __WATER_IN__,                    !- Heat Recovery Water Inlet Node Name
    __WATER_OUT__,                   !- Heat Recovery Water Outlet Node Name
    0.0004,                          !- Heat Recovery Water Maximum Flow Rate {m3/s}
    __FC__ FC Exhaust Air Outlet Node, !- Exhaust Outlet Air Node Name
    CONDENSING,                      !- Heat Exchanger Calculation Method
    ,                                !- Method 1 Heat Exchanger Effectiveness
    83.1,                            !- Method 2 Parameter hxs0
    4798,                            !- Method 2 Parameter hxs1
    -138E+3,                         !- Method 2 Parameter hxs2
    -353.8E+3,                       !- Method 2 Parameter hxs3
    5.15E+8,                         !- Method 2 Parameter hxs4
    ,                                !- Method 3 h0Gas Coefficient
    ,                                !- Method 3 NdotGasRef Coefficient
    ,                                !- Method 3 n Coefficient
    ,                                !- Method 3 Gas Area {m2}
    ,                                !- Method 3 h0 Water Coefficient
    ,                                !- Method 3 N dot Water ref Coefficient
    ,                                !- Method 3 m Coefficient
    ,                                !- Method 3 Water Area {m2}
    ,                                !- Method 3 F Adjustment Factor
    -1.96E-4,                        !- Method 4 hxl1 Coefficient
    3.1E-3,                          !- Method 4 hxl2 Coefficient
    35.0;                            !- Method 4 Condensation Threshold {C}

Generator:FuelCell:ElectricalStorage,
    __FC__ FC Battery,               !- Name
    SimpleEfficiencyWithConstraints, !- Choice of Model
    1.0,                             !- Nominal Charging Energetic Efficiency
    1.0,                             !- Nominal Discharging Energetic Efficiency
    0,                               !- Simple Maximum Capacity {J}
    0,                               !- Simple Maximum Power Draw {W}
    0,                               !- Simple Maximum Power Store {W}
    0;                               !- Initial Charge State {J}

Generator:FuelCell:Inverter,
    __FC__ FC Inverter,              !- Name
    QUADRATIC,                       !- Inverter Efficiency Calculation Mode
    ,                                !- Inverter Efficiency
    __FC__ FC Inverter Quadratic;    !- Efficiency Function of DC Power Curve Name

Curve:Quadratic,
    __FC__ FC Inverter Quadratic,    !- Name
    0.560717,                        !- Coefficient1 Constant
    1.24019E-4,                      !- Coefficient2 x
    -2.01648E-8,                     !- Coefficient3 x**2
    -1.0E+10,                        !- Minimum Value of x
    1.0E+10;                         !- Maximum Value of x

OutdoorAir:NodeList,
    __FC__ FC Air Inlet Node,        !- Node 1 (fuel-cell air / fuel / water temperature reference)
    __FC__ FC Dilution Inlet Node;   !- Node 2 (dilution air inlet)
";

            // Route skin/auxiliary losses to ambient when no zone is given, or to the named zone.
            string powerModuleZone;
            string auxiliarySkinDestination;
            string auxiliaryZone;
            if (string.IsNullOrEmpty(zoneName))
            {
                powerModuleZone = "";                          // blank Zone Name -> losses to ambient
                auxiliarySkinDestination = "AirInletForFuelCell";
                auxiliaryZone = "";
            }
            else
            {
                powerModuleZone = zoneName;
                auxiliarySkinDestination = "SurroundingZone";
                auxiliaryZone = zoneName;
            }

            string result = template;
            result = result.Replace("__WATER_IN__", waterInletNode);
            result = result.Replace("__WATER_OUT__", waterOutletNode);
            result = result.Replace("__POWERMODULE_ZONE__", powerModuleZone);
            result = result.Replace("__AUX_SKIN_DEST__", auxiliarySkinDestination);
            result = result.Replace("__AUX_ZONE__", auxiliaryZone);
            result = result.Replace("__FC__", generatorName);
            return result;
        }

        // Returns the object of the given type whose name matches objectName, or throws.
        private IdfObject FindByName(string objectType, string objectName)
        {
            foreach (IdfObject idfObject in idfReader[objectType])
            {
                if (idfObject[0].Value == objectName)
                {
                    return idfObject;
                }
            }
            throw new MissingFieldException("Cannot find object named '" + objectName + "' of type: " + objectType);
        }

        // Finds the Branch that carries the given component (type + name) and returns its
        // component inlet/outlet node names as { inlet, outlet }.
        private string[] GetBranchComponentNodes(string componentType, string componentName)
        {
            foreach (IdfObject branch in idfReader["Branch"])
            {
                for (int i = 0; i < (branch.Count - 1); i++)
                {
                    if (branch[i].Value == componentType && branch[i + 1].Value == componentName)
                    {
                        string[] nodes = new string[2];
                        nodes[0] = branch[i + 2].Value; // Component inlet node name
                        nodes[1] = branch[i + 3].Value; // Component outlet node name
                        return nodes;
                    }
                }
            }
            throw new MissingFieldException("Could not locate the Branch carrying component: " + componentName);
        }

        // Replaces an adjacent (object type, object name) pair inside a list object
        // (e.g. Branch, PlantEquipmentList). Type comes first, then name.
        private void ReplaceObjectTypeInList(
            string listName,
            string oldObjectType,
            string oldObjectName,
            string newObjectType,
            string newObjectName)
        {
            var idfObjects = idfReader[listName];

            bool replacementMade = false;

            foreach (IdfObject idfObject in idfObjects)
            {
                if (replacementMade)
                {
                    break;
                }

                for (int i = 0; i < (idfObject.Count - 1); i++)
                {
                    Field currentField = idfObject[i];
                    Field nextField = idfObject[i + 1];

                    // Note: comparison is case-sensitive here.
                    if (currentField.Value == oldObjectType && nextField.Value == oldObjectName)
                    {
                        currentField.Value = newObjectType;
                        nextField.Value = newObjectName;
                        replacementMade = true;
                        break;
                    }
                }
            }
        }

        // In a name-then-type list (e.g. ElectricLoadCenter:Generators), finds the entry whose
        // name field equals generatorName and whose immediately following type field equals
        // oldType, and changes only that type field to newType. Scoping by name leaves other
        // generators in the same list untouched.
        private void ReplaceTypeForNamedGenerator(string listName, string generatorName, string oldType, string newType)
        {
            foreach (IdfObject idfObject in idfReader[listName])
            {
                for (int i = 0; i < (idfObject.Count - 1); i++)
                {
                    Field nameField = idfObject[i];
                    Field typeField = idfObject[i + 1];

                    if (nameField.Value == generatorName && typeField.Value == oldType)
                    {
                        typeField.Value = newType;
                        return;
                    }
                }
            }
        }
    }
}