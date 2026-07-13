/*
 * Monodraught HVR APX — DesignBuilder EMS Control Script
 *
 * WHAT THIS SCRIPT REPRESENTS
 * Models one or more Monodraught HVR APX heat recovery ventilation units serving a zone,
 * together with their integrated louvre vent(s). At each HVAC timestep it:
 *   - Enables mechanical ventilation with heat recovery (50% sensible effectiveness),
 *     sized to 0.22 m3/s (180 l/s) per unit.
 *   - Opens the integrated louvre(s) for natural ventilation when outdoor temperature is
 *     above 15C and the zone is warm or CO2-laden enough to need it.
 *   - Otherwise drives the mechanical fan through 5 discrete speed steps chosen from zone
 *     temperature (or CO2 concentration once the temperature setpoint is exceeded),
 *     halving the flow rate whenever outdoor air is cold.
 *   - Estimates unit fan electrical power from a constant specific fan power (SFP = 0.2 W per l/s).
 *   - Scales design flow rate, fan speeds, and expected vent count with the number of HVR
 *     units declared on the zone's tag, so a single zone can host multiple units.
 *
 * QUICK SETUP GUIDE
 *   1. Site level: enable CO2 concentration modelling (Site Details, Outdoor Air CO2 and
 *      Contaminants) and set Natural Ventilation Method to Calculated Natural Ventilation.
 *   2. Tag each HVR zone HVRUNIT for 1 unit, or HVRUNIT followed by a number for more,
 *      e.g. HVRUNIT3 for 3 units.
 *   3. For every unit, draw a Vent (0.9 m width x 0.3 m height, discharge coefficient 0.12)
 *      tagged HVRVENT. A zone tagged HVRUNIT3 needs 3 such vents.
 *   4. Paste this script into Tools, Scripts, CS-Script, then click Enable Program.
 *
 * Revision 3: Added support for multiple HVR units per zone. The zone's "HVRUNIT" tag may now
 *             carry a numeric suffix (e.g. "HVRUNIT3") declaring how many units serve the zone.
 *             That count scales MechanicalVentilationDesignFlowRate and each fan-speed flow rate,
 *             and the zone must contain exactly that many HVRVENT-tagged vents (one Venting_Opening_Factor
 *             actuator per vent). Fan power scales automatically since it is derived from total zone flow.
 *             Fixed a warm-branch EMS bug: the fan-speed gate compared zone temperature against
 *             Speed_1_Flow_rate (a flow rate, not a temperature) instead of Temp_Threshold_1, so it
 *             was effectively always true. It now correctly gates on Temp_Threshold_1.
 *
 * Revision 2: Excluded zones are now disregarded - zone selection skips any zone whose
 *             "IncludeZone" attribute is "0", so settings and EMS objects are no longer
 *             generated for zones left out of the simulation. Removed unused GetHrvZones().
 *             Vent dimensions are now read via the Opening.Width / Opening.Height
 *             properties (doubles, in metres) instead of parsing the OpeningWidth /
 *             OpeningHeight string attributes.
 *
 * Revision 1: Added zone equipment load to represent HVR unit fan power consumption (SFP = 0.2 W/(l/s)).
 *             Fan power is an estimate based on a constant SFP; verify against manufacturer data for accuracy.
 *
 * This script is provided as-is to assist with Monodraught HVR APX unit control within DesignBuilder.
 * DesignBuilder Software Ltd does not accept any responsibility or liability for the accuracy,
 * suitability, or results obtained from using this script. Users should independently verify and
 * validate the implementation and simulation outcomes in their own modeling context.
 */

using DB.Api;
using DB.Extensibility.Contracts;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime;
using System.Text;
using System.Windows.Forms;

namespace DB.Extensibility.Scripts
{
    public class ApplyHvrControl : ScriptBase, IScript
    {
        private const string ZoneTag = "HVRUNIT";
        private const string VentTag = "HVRVENT";
        private const double Sfp = 0.2; // W/(l/s)
        private const decimal DesignFlowRatePerUnit = 0.22m; // m3/s per single HVR unit
        private string[] InvalidEmsChars = new[] { " ", "-", ":", "." };

        // Resolved per zone once per run: unit count (from the zone tag suffix) + validated vents.
        private class HvrZoneInfo
        {
            public Zone Zone;
            public int NumUnits;
            public List<Opening> Vents;
        }

        public override void BeforeEnergySimulation()
        {
            List<Zone> zones = ApplyHvrSettings();
            if (zones.Count == 0)
            {
                throw new Exception(string.Format("No zones with the '{0}' tag were found in the model. HRV control was not applied.", ZoneTag));
            }
            string code = GenerateHrvEmsCode(zones);
            File.AppendAllText(ApiEnvironment.EnergyPlusInputIdfPath, code);
        }

        private string GenerateHrvEmsCode(List<Zone> zones)
        {
            // Resolve unit count + vents once per zone; reused across both EMS passes below.
            List<HvrZoneInfo> zoneInfos = zones.Select(zone =>
            {
                int numUnits = GetHvrUnitCount(zone);
                List<Opening> vents = GetHrvVents(zone, numUnits);
                return new HvrZoneInfo { Zone = zone, NumUnits = numUnits, Vents = vents };
            }).ToList();

            // Generate EMS objects for HRV control
            string sharedEmsCode = GetSharedEmsCode();
            StringBuilder emsContent = new StringBuilder(sharedEmsCode);

            foreach (var info in zoneInfos)
            {
                string zoneEmsCode = GetZoneEmsCode(info.Zone, info.Vents);
                emsContent.AppendLine(zoneEmsCode);
            }
            // Generate EMS program for HRV control logic
            string sharedProgramCode = GetGetSharedEmsProgramCode();
            emsContent.AppendLine(sharedProgramCode);
            foreach (var info in zoneInfos)
            {
                string zoneEmsCode = GetZoneEmsProgramCode(info.Zone, info.Vents, info.NumUnits);
                emsContent.AppendLine(zoneEmsCode);
            }
            // Apply a proper object termination
            emsContent[emsContent.Length - 1] = ';';

            string zoneSummary = string.Join(", ", zoneInfos.Select(info =>
                string.Format("{0} ({1} unit{2})", info.Zone.GetAttribute("Title"), info.NumUnits, info.NumUnits == 1 ? "" : "s")));
            MessageBox.Show(string.Format("HVR unit control applied in zones: {0}.", zoneSummary));

            return emsContent.ToString();
        }

        private string GetVariableName(string idfName)
        {
            foreach (string character in InvalidEmsChars)
            {
                idfName = idfName.Replace(character, "_");
            }
            return idfName;
        }

        private string GetItemFromTable(string tableName, string fieldName, int handle)
        {
            Site site = ApiEnvironment.Site;
            Table table = site.GetTable(tableName);
            Record record = table.Records.GetRecordFromHandle(handle);
            return record[fieldName];
        }

        // Parses the number of HVR units from the zone's ObjectTag, e.g. "HVRUNIT" => 1, "HVRUNIT3" => 3.
        private int GetHvrUnitCount(Zone zone)
        {
            string tag = zone.GetAttribute("ObjectTag") ?? string.Empty;
            string suffix = tag.Substring(ZoneTag.Length);
            if (string.IsNullOrEmpty(suffix))
            {
                return 1;
            }
            string invalidTagMessage = string.Format(
                "Zone '{0}' has an unrecognised HVR tag '{1}'. Expected '{2}' for 1 unit, or '{2}' followed by a positive integer (e.g. '{2}3') for multiple units.",
                zone.GetAttribute("Title"), tag, ZoneTag);

            int unitCount;
            if (!int.TryParse(suffix, NumberStyles.Integer, CultureInfo.InvariantCulture, out unitCount))
            {
                throw new Exception(invalidTagMessage);
            }
            if (unitCount < 1)
            {
                throw new Exception(invalidTagMessage);
            }
            return unitCount;
        }

        private string GetZoneEmsCode(Zone zone, List<Opening> vents)
        {
            string zoneVariableName = GetVariableName(zone.IdfName);
            string scheduleName = GetItemFromTable("Schedules", "Name", zone.GetAttributeAsInt("MechanicalVentilationSchedule"));

            // Zone-level sensors/actuators: one ideal-loads system and one fan-power meter per zone,
            // shared across however many HVR units (vents) serve it.
            string zoneTemplate = @"
! Read zone temperature and CO2 levels
EnergyManagementSystem:Sensor,
   Zone_Air_CO2_Concentration_{1},
   {0},
   Zone Air CO2 Concentration;

EnergyManagementSystem:Sensor,
   Zone_Mean_Air_Temperature_{1},
   {0},
   Zone Mean Air Temperature;

! Read mechanical ventilation availability
EnergyManagementSystem:Sensor,
   Zone_Mech_Vent_Availability_{1},
   {2},
   Schedule Value;

! Control OA flow rate ({3} unit(s) served by this zone's ideal loads system)
EnergyManagementSystem:Actuator,
   Outdoor_Air_Mass_Flow_Rate_{1},
   {0} IDEAL LOADS AIR,
   Ideal Loads Air System,
   Outdoor Air Mass Flow Rate;

! Zone equipment load representing combined HVR unit fan power
OtherEquipment,
   HVR Unit {0} Electricity Rate,
   Electricity,
   {0},
   ON,
   EquipmentLevel,
   1,
   ,
   ,
   0,
   0,
   0.5;

EnergyManagementSystem:Actuator,
   HVR_Unit_Power_{1},
   HVR Unit {0} Electricity Rate,
   OtherEquipment,
   Power Level;
";
            StringBuilder code = new StringBuilder(string.Format(
                zoneTemplate, zone.IdfName, zoneVariableName, scheduleName, vents.Count));

            // Per-vent actuator: one integrated louvre actuator per physical HVR unit in the zone.
            string ventTemplate = @"
! Control integrated louvre ({0})
EnergyManagementSystem:Actuator,
   Venting_Opening_Factor_{2},
   {1},
   AirFlow Network Window/Door Opening,
   Venting Opening Factor;
";
            foreach (Opening vent in vents)
            {
                string ventIdfName = vent.GetAttribute("SSEPObjectNameInOP");
                string ventVariableName = GetVariableName(ventIdfName);
                code.Append(string.Format(ventTemplate, vent.GetAttribute("Title"), ventIdfName, ventVariableName));
            }

            return code.ToString();
        }

        private string GetZoneEmsProgramCode(Zone zone, List<Opening> vents, int numUnits)
        {
            string zoneVariableName = GetVariableName(zone.IdfName);
            List<string> ventVariableNames = vents
                .Select(vent => GetVariableName(vent.GetAttribute("SSEPObjectNameInOP")))
                .ToList();

            // One SET line per vent so all louvres in the zone move together.
            string ventOpenText = string.Join("\n         ", ventVariableNames.Select(name =>
                string.Format("SET Venting_Opening_Factor_{0} = 1,", name)));
            string ventCloseText = string.Join("\n         ", ventVariableNames.Select(name =>
                string.Format("SET Venting_Opening_Factor_{0} = 0,", name)));

            string template = @"
   ! Natural ventilation is only enabled when OA temperature is greater than 15C
   IF Site_Outdoor_Air_Drybulb_Temperature > 15,
      IF Zone_Mean_Air_Temperature_{1} > Zone_Natural_Ventilation_Temp_Setpoint,
         ! Zone temperature exceeds setpoint
         {2}
      ELSEIF Zone_Air_CO2_Concentration_{1} > Zone_Natural_Ventilation_CO2_Setpoint && Zone_Mean_Air_Temperature_{1} > 16,
         ! Zone CO2 exceeds setpoint and internal temperature is above minimum threshold
         {2}
      ELSE,
         {3}
      ENDIF,
   ELSE,
      {3}
   ENDIF,

   IF Zone_Mech_Vent_Availability_{1} == 0,
      ! Disable mechanical ventilation when schedule is off
      SET Outdoor_Air_Mass_Flow_Rate_{1} = Null,
      SET HVR_Unit_Power_{1} = 0,
      RETURN,
   ENDIF,

   IF Site_Outdoor_Air_Drybulb_Temperature > 15,
      IF Zone_Mean_Air_Temperature_{1} > Temp_Threshold_1,
         ! Zone temperature exceeds the minimum temperature setpoint, find the fan speed
         IF Zone_Mean_Air_Temperature_{1} > Temp_Threshold_5,
            SET Flow_Rate = Speed_5_Flow_rate * {0},
         ELSEIF Zone_Mean_Air_Temperature_{1} > Temp_Threshold_4,
            SET Flow_Rate = Speed_4_Flow_rate * {0},
         ELSEIF Zone_Mean_Air_Temperature_{1} > Temp_Threshold_3,
            SET Flow_Rate = Speed_3_Flow_rate * {0},
         ELSEIF Zone_Mean_Air_Temperature_{1} > Temp_Threshold_2,
            SET Flow_Rate = Speed_2_Flow_rate * {0},
         ELSEIF Zone_Mean_Air_Temperature_{1} > Temp_Threshold_1,
            SET Flow_Rate = Speed_1_Flow_rate * {0},
         ELSE,
            SET Flow_Rate = Unit_Disabled,
         ENDIF,
      ELSEIF Zone_Air_CO2_Concentration_{1} > CO2_Threshold_1 && Zone_Mean_Air_Temperature_{1} > 16,
         ! Zone CO2 concentration exceeds the minimum CO2 setpoint, find the fan speed
         IF Zone_Air_CO2_Concentration_{1} > CO2_Threshold_5,
            SET Flow_Rate = Speed_5_Flow_rate * {0},
         ELSEIF Zone_Air_CO2_Concentration_{1} > CO2_Threshold_4,
            SET Flow_Rate = Speed_4_Flow_rate * {0},
         ELSEIF Zone_Air_CO2_Concentration_{1} > CO2_Threshold_3,
            SET Flow_Rate = Speed_3_Flow_rate * {0},
         ELSEIF Zone_Air_CO2_Concentration_{1} > CO2_Threshold_2,
            SET Flow_Rate = Speed_2_Flow_rate * {0},
         ELSEIF Zone_Air_CO2_Concentration_{1} > CO2_Threshold_1,
            SET Flow_Rate = Speed_1_Flow_rate * {0},
         ELSE,
            SET Flow_Rate = Unit_Disabled,
         ENDIF,
      ELSE,
         SET Flow_Rate = Unit_Disabled,
      ENDIF,
      SET Outdoor_Air_Mass_Flow_Rate_{1} = Flow_Rate,

   ! Define rules for cold conditions
   ELSE,
      IF Zone_Mean_Air_Temperature_{1} > Cold_Temp_Threshold_1 && Zone_Air_CO2_Concentration_{1} < CO2_Threshold_1,
         IF Zone_Mean_Air_Temperature_{1} > Cold_Temp_Threshold_5,
            SET Flow_Rate = Speed_5_Flow_rate * {0},
         ELSEIF Zone_Mean_Air_Temperature_{1} > Cold_Temp_Threshold_4,
            SET Flow_Rate = Speed_4_Flow_rate * {0},
         ELSEIF Zone_Mean_Air_Temperature_{1} > Cold_Temp_Threshold_3,
            SET Flow_Rate = Speed_3_Flow_rate * {0},
         ELSEIF Zone_Mean_Air_Temperature_{1} > Cold_Temp_Threshold_2,
            SET Flow_Rate = Speed_2_Flow_rate * {0},
         ELSEIF Zone_Mean_Air_Temperature_{1} > Cold_Temp_Threshold_1,
            SET Flow_Rate = Speed_1_Flow_rate * {0},
         ELSE,
            SET Flow_Rate = Unit_Disabled,
         ENDIF,
      ELSEIF Site_Outdoor_Air_Drybulb_Temperature > 10 && Zone_Mean_Air_Temperature_{1} > 16 && Zone_Air_CO2_Concentration_{1} > CO2_Threshold_1,
         IF Zone_Air_CO2_Concentration_{1} > CO2_Threshold_5,
            SET Flow_Rate = Speed_5_Flow_rate * {0},
         ELSEIF Zone_Air_CO2_Concentration_{1} > CO2_Threshold_4,
            SET Flow_Rate = Speed_4_Flow_rate * {0},
         ELSEIF Zone_Air_CO2_Concentration_{1} > CO2_Threshold_3,
            SET Flow_Rate = Speed_3_Flow_rate * {0},
         ELSEIF Zone_Air_CO2_Concentration_{1} > CO2_Threshold_2,
            SET Flow_Rate = Speed_2_Flow_rate * {0},
         ELSEIF Zone_Air_CO2_Concentration_{1} > CO2_Threshold_1,
            SET Flow_Rate = Speed_1_Flow_rate * {0},
         ELSE,
            SET Flow_Rate = Unit_Disabled,
         ENDIF,
      ELSE,
         SET Flow_Rate = Unit_Disabled,
      ENDIF,
      ! Reduce the ventilation rate by half in cold conditions
      SET Outdoor_Air_Mass_Flow_Rate_{1} = Flow_Rate / 2,
   ENDIF,
   SET HVR_Unit_Power_{1} = Outdoor_Air_Mass_Flow_Rate_{1} * 1000 / 1.2 * {4},";
            return string.Format(template, numUnits, zoneVariableName, ventOpenText, ventCloseText, Sfp);
        }

        private string GetSharedEmsCode()
        {
            return @"
! Read outdoor air temperature
EnergyManagementSystem:Sensor,
   Site_Outdoor_Air_Drybulb_Temperature,
   Environment,
   Site Outdoor Air Drybulb Temperature;

! Initialize variables
EnergyManagementSystem:GlobalVariable, Temp_Threshold_1;
EnergyManagementSystem:GlobalVariable, Temp_Threshold_2;
EnergyManagementSystem:GlobalVariable, Temp_Threshold_3;
EnergyManagementSystem:GlobalVariable, Temp_Threshold_4;
EnergyManagementSystem:GlobalVariable, Temp_Threshold_5;

EnergyManagementSystem:GlobalVariable, Cold_Temp_Threshold_1;
EnergyManagementSystem:GlobalVariable, Cold_Temp_Threshold_2;
EnergyManagementSystem:GlobalVariable, Cold_Temp_Threshold_3;
EnergyManagementSystem:GlobalVariable, Cold_Temp_Threshold_4;
EnergyManagementSystem:GlobalVariable, Cold_Temp_Threshold_5;

EnergyManagementSystem:GlobalVariable, CO2_Threshold_1;
EnergyManagementSystem:GlobalVariable, CO2_Threshold_2;
EnergyManagementSystem:GlobalVariable, CO2_Threshold_3;
EnergyManagementSystem:GlobalVariable, CO2_Threshold_4;
EnergyManagementSystem:GlobalVariable, CO2_Threshold_5;

! Per-unit fan speed flow rates; scaled by each zone's HVR unit count (from its zone tag) at the point of use
EnergyManagementSystem:GlobalVariable, Speed_1_Flow_Rate;
EnergyManagementSystem:GlobalVariable, Speed_2_Flow_Rate;
EnergyManagementSystem:GlobalVariable, Speed_3_Flow_Rate;
EnergyManagementSystem:GlobalVariable, Speed_4_Flow_Rate;
EnergyManagementSystem:GlobalVariable, Speed_5_Flow_Rate;

EnergyManagementSystem:GlobalVariable, Unit_Disabled;

EnergyManagementSystem:ProgramCallingManager,
   BeginEnvCaller,
   BeginNewEnvironment,
   InitVariables;

EnergyManagementSystem:Program,
   InitVariables,
   ! Define stepped temperature thresholds
   SET Temp_Threshold_1 = 23,
   SET Temp_Threshold_2 = 24,
   SET Temp_Threshold_3 = 25,
   SET Temp_Threshold_4 = 26,
   SET Temp_Threshold_5 = 28,

   ! Define stepped temperature thresholds
   SET Cold_Temp_Threshold_1 = 25,
   SET Cold_Temp_Threshold_2 = 26,
   SET Cold_Temp_Threshold_3 = 27,
   SET Cold_Temp_Threshold_4 = 28,
   SET Cold_Temp_Threshold_5 = 30,

   ! Define stepped CO2 ppm thresholds
   SET CO2_Threshold_1 = 850,
   SET CO2_Threshold_2 = 900,
   SET CO2_Threshold_3 = 950,
   SET CO2_Threshold_4 = 1000,
   SET CO2_Threshold_5 = 1050,

   ! Define discrete per-unit fan mass flow rates
   SET Unit_Disabled = 0,
   SET Speed_1_Flow_rate = 80 * 1.2 / 1000,
   SET Speed_2_Flow_rate = 105 * 1.2 / 1000,
   SET Speed_3_Flow_rate = 130 * 1.2 / 1000,
   SET Speed_4_Flow_rate = 160 * 1.2 / 1000,
   SET Speed_5_Flow_rate = 180 * 1.2 / 1000;
";
        }


        private string GetGetSharedEmsProgramCode()
        {
            return @"
  EnergyManagementSystem:ProgramCallingManager,
   Caller,
   InsideHVACSystemIterationLoop,
   ControlHVRUnit;

EnergyManagementSystem:Program,
   ControlHVRUnit,
   ! Define rules for natural ventilation mode
   SET Zone_Natural_Ventilation_Temp_Setpoint = 19,       ! range between 19.0 and 22.0C
   SET Zone_Natural_Ventilation_CO2_Setpoint = 800,       ! range between 800 and 900ppm";
        }

        private List<Zone> ApplyHvrSettings()
        {
            Site site = ApiEnvironment.Site;
            Building building = site.Buildings[ApiEnvironment.CurrentBuildingIndex];
            List<Zone> zones = building.BuildingBlocks
                .SelectMany(block => block.Zones)
                .Where(zone => (zone.GetAttribute("ObjectTag") ?? string.Empty).StartsWith(ZoneTag, StringComparison.Ordinal))
                .Where(zone => !zone.IsChildZone)
                .Where(zone => !zone.GetAttribute("IncludeZone").Equals("0"))
                .ToList();

            foreach (Zone zone in zones)
            {
                int numUnits = GetHvrUnitCount(zone);
                decimal designFlowRate = DesignFlowRatePerUnit * numUnits;

                zone.SetAttribute("MechanicalVentilationOn", "1");
                zone.SetAttribute("MechanicalVentilationRateType", "6");
                zone.SetAttribute("MechanicalVentilationDesignFlowRate", designFlowRate.ToString(CultureInfo.InvariantCulture));
                zone.SetAttribute("HeatRecoveryOn", "1");
                zone.SetAttribute("HeatRecoveryType", "1");
                zone.SetAttribute("SensibleHeatRecoveryEffectiveness", "0.50");

                int mechVentEnabled = zone.GetAttributeAsInt("MechanicalVentilationOn");
                if (mechVentEnabled == 0)
                {
                    throw new Exception(string.Format("Zone '{0}' has mechanical ventilation disabled. Please enable it to use HRV control.", zone.GetAttribute("Title")));
                }
            }
            return zones;
        }

        // Finds every HVRVENT-tagged vent in the zone and validates dimensions, discharge coefficient,
        // and that the vent count matches the unit count declared in the zone's tag.
        private List<Opening> GetHrvVents(Zone zone, int expectedUnitCount)
        {
            List<Opening> vents = zone.Surfaces
                .SelectMany(surface => surface.Adjacencies)
                .SelectMany(adjacency => adjacency.Openings)
                .Where(opening => opening.Type == OpeningType.Vent && opening.GetAttribute("ObjectTag").Equals(VentTag))
                .ToList();

            if (vents.Count == 0)
            {
                throw new Exception(string.Format("Zone '{0}' does not include a vent with the '{1}' tag.", zone.GetAttribute("Title"), VentTag));
            }
            if (vents.Count != expectedUnitCount)
            {
                throw new Exception(string.Format(
                    "Zone '{0}' is tagged for {1} HVR unit(s) but has {2} vent(s) tagged '{3}'. Place exactly {1} vent(s) in the zone or correct the zone tag.",
                    zone.GetAttribute("Title"), expectedUnitCount, vents.Count, VentTag));
            }

            foreach (Opening vent in vents)
            {
                // Validate vent dimensions (Width/Height properties return metres as doubles)
                double ventWidth = vent.Width;
                double ventHeight = vent.Height;
                if (!(ventWidth > 0.89 && ventWidth < 0.91) && !(ventHeight > 0.29 && ventHeight < 0.31))
                {
                    throw new Exception(string.Format("Vent '{0}' in zone '{1}' are out of bounds. The expected size is 0.9x0.3m.", vent.GetAttribute("Title"), zone.GetAttribute("Title")));
                }
                // Validate vent discharge coefficient
                int ventId = vent.GetAttributeAsInt("VentType");

                string item = GetItemFromTable("Vents", "CD", ventId);
                double dischargeCoefficient = double.Parse(item, CultureInfo.InvariantCulture);
                if (dischargeCoefficient != 0.12)
                {
                    throw new Exception(string.Format("Vent '{0}' in zone '{1}' has an invalid discharge coefficient: {2}. The expected value is 0.12.", vent.GetAttribute("Title"), zone.GetAttribute("Title"), dischargeCoefficient));
                }
            }

            return vents;
        }
    }
}