/*
Custom Meters Script

Purpose:
This DesignBuilder C# script defines aggregated EnergyPlus Meter:Custom and Output:Meter objects
for DesignBuilder models, splitting results into residential and non-residential end uses.

Main Steps:
1) Inspect zone activity templates and equipment attributes in the DesignBuilder data model
2) Classify each zone as residential / non-residential using its Activity Template name
3) Build Meter:Custom definitions by aggregating standard EnergyPlus zone end-use meter names
    * Interior lighting electricity
    * Misc/process/catering gas equipment
    * Misc/process/catering and other pure-electric equipment
4) Inject Meter:Custom and Output:Meter objects into the IDF before the EnergyPlus simulation runs
     - ResidentialLights
     - Non-ResidentialLights
     - ResidentialGas
     - Non-ResidentialGas
     - ResidentialElectricEquipment
     - Non-ResidentialElectricEquipment

How to Use:

Configuration
- Confirm that the Activity Template names in _residentialActivityTemplateNames match those used in your project templates, and adjust the list if needed.
- Reporting frequencies for Output:Meter are set in CustomMeter.ToIdfString() via the reportingFrequencies parameter (defaults to Hourly/Daily/Monthly/RunPeriod).

Notes:
- Only non-child zones are considered
- A residential/non-residential meter is only created if at least one qualifying zone contributes to it (empty meters are not written).
- Update the resource type strings or meter name patterns if your IDF uses non-standard naming conventions.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. 
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using DB.Api;
using DB.Extensibility.Contracts;
using EpNet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DB.Extensibility.Scripts
{
    public class CustomMeters : ScriptBase, IScript
    {
        private IdfReader idfReader; // Reader used to load/modify the EnergyPlus IDF

        // Activity Template names treated as "residential"
        private readonly List<string> _residentialActivityTemplateNames = new List<string>
        {
            "Domestic Bedroom",
            "Domestic Bathroom",
            "Domestic Kitchen"
        };

        // Attribute / table / column names in the DB data model
        private const string ActivityTemplateIdAttr = "ActivityTemplateId";
        private const string ActivityTemplateTableName = "ActivityTemplates";
        private const string TemplateNameColumn = "Name";
        private const string LightingOnAttr = "FluorescentLightingOn";
        private const string IdfNameAttr = "SSEPObjectNameInOP";

        // Zone attributes for purely electric equipment switches
        private readonly string[] PureElectricEquipmentAttrs = new string[]
        {
            "EquipmentOn",
            "ComputersOn"
        };

        // Zone attributes for misc equipment: [on-flag, fuel-type]
        private readonly string[,] MiscEquipmentAttrs = new string[,]
        {
            { "MiscellaneousOn", "MiscellaneousFuel" },
            { "CateringOn",      "CateringFuel" },
            { "ProcessOn",       "ProcessFuel" }
        };

        // Helper class representing one EnergyPlus Meter:Custom + Output:Meter set
        public class CustomMeter
        {
            public string Name;          // Meter name used in IDF
            public string ResourceType;  // EnergyPlus resource type (e.g. Electricity, NaturalGas)

            // (Key Name, Variable or Meter Name) pairs for Meter:Custom
            private readonly List<Tuple<string, string>> _keyValuePairs = new List<Tuple<string, string>>();

            public CustomMeter(string name, string resourceType = "Electricity")
            {
                Name = name;
                ResourceType = resourceType;
            }

            // Add one Output:Variable entry to the custom meter
            public void AddVariable(string keyName, string variableName)
            {
                _keyValuePairs.Add(new Tuple<string, string>(keyName, variableName));
            }

            // Add one existing meter entry to the custom meter (no key name)
            public void AddMeter(string meterName)
            {
                _keyValuePairs.Add(new Tuple<string, string>(string.Empty, meterName));
            }

            // True if there are no entries in this custom meter
            public bool IsEmpty()
            {
                return _keyValuePairs.Count == 0;
            }

            // Build Meter:Custom and corresponding Output:Meter objects as IDF text
            public string ToIdfString(string[] reportingFrequencies = null)
            {
                // Default reporting frequencies if none provided
                if (reportingFrequencies == null)
                {
                    reportingFrequencies = new string[] { "Hourly", "Daily", "Monthly", "RunPeriod" };
                }

                var idfText = new StringBuilder();

                // Meter:Custom object header
                idfText.AppendLine("Meter:Custom,");
                idfText.AppendLine("  " + Name + ",                       !- Name");
                idfText.AppendLine("  " + ResourceType + ",              !- Resource Type");

                // Add each (key, variable/meter) pair
                for (int i = 0; i < _keyValuePairs.Count; i++)
                {
                    var entry = _keyValuePairs[i];
                    bool isLast = i == _keyValuePairs.Count - 1;

                    idfText.AppendLine("  " + entry.Item1 + ",                   !- Key Name " + (i + 1));
                    idfText.Append("  " + entry.Item2); // Output:Variable or Meter Name

                    if (isLast)
                        idfText.AppendLine(";                              !- Output Variable or Meter Name " + (i + 1));
                    else
                        idfText.AppendLine(",                              !- Output Variable or Meter Name " + (i + 1));
                }

                idfText.AppendLine();

                // Output:Meter objects for each reporting frequency
                foreach (string interval in reportingFrequencies)
                {
                    idfText.AppendLine(string.Format("Output:Meter, {0}, {1};", Name, interval));
                }

                idfText.AppendLine();
                return idfText.ToString();
            }
        }

        // Entry point before EnergyPlus simulation: create meters and save IDF
        public override void BeforeEnergySimulation()
        {
            Site site = ApiEnvironment.Site;  // Current DB site
            Building building = site.Buildings[ApiEnvironment.CurrentBuildingIndex]; // Active building

            // Load IDF/IDD into EpNet reader
            idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // Build all custom meters and inject them into the IDF
            CreateCustomMeters(building);

            // Save modified IDF back to disk
            idfReader.Save();
        }

        // Check if a given zone belongs to a residential Activity Template
        private bool IsZoneResidential(Zone zone)
        {
            int templateId = zone.GetAttributeAsInt(ActivityTemplateIdAttr);

            Site site = ApiEnvironment.Site;
            Table table = site.GetTable(ActivityTemplateTableName);
            Record record = table.Records.GetRecordFromHandle(templateId);

            // Template name must be one of the configured residential template names
            return _residentialActivityTemplateNames.Contains(record[TemplateNameColumn]);
        }

        // Split all zones into residential / non-residential lists
        private void GetResidentialAndNonResidentialZones(
            Building building,
            out List<Zone> residentialZones,
            out List<Zone> nonResidentialZones)
        {
            residentialZones = new List<Zone>();
            nonResidentialZones = new List<Zone>();

            foreach (BuildingBlock block in building.BuildingBlocks)
            {
                foreach (Zone zone in block.Zones)
                {
                    if (zone.IsChildZone)
                        continue; // Skip child zones; only top-level zones count

                    if (IsZoneResidential(zone))
                    {
                        residentialZones.Add(zone);
                    }
                    else
                    {
                        nonResidentialZones.Add(zone);
                    }
                }
            }
        }

        // Build a custom meter that sums lighting electricity for the given zones
        private CustomMeter GetLightingMeter(string name, List<Zone> zones)
        {
            CustomMeter meter = new CustomMeter(name);

            foreach (Zone zone in zones)
            {
                // Only include zones where lighting is enabled 
                if (zone.GetAttributeAsInt(LightingOnAttr) == 1)
                {
                    string zoneIdfName = zone.GetAttribute(IdfNameAttr);
                    string lightingMeterName = "InteriorLights:Electricity:Zone:" + zoneIdfName;
                    meter.AddMeter(lightingMeterName);
                }
            }

            return meter;
        }

        // Build a custom meter that sums gas equipment loads for the given zones
        private CustomMeter GetEquipmentGasMeter(string name, List<Zone> zones)
        {
            CustomMeter meter = new CustomMeter(name, "NaturalGas");

            foreach (Zone zone in zones)
            {
                // Loop over misc equipment types (misc/catering/process)
                for (int i = 0; i < MiscEquipmentAttrs.GetLength(0); i++) // rows
                {
                    string onAttr = MiscEquipmentAttrs[i, 0];
                    string fuelAttr = MiscEquipmentAttrs[i, 1];

                    int equipmentOn = zone.GetAttributeAsInt(onAttr);
                    string fuel = zone.GetAttribute(fuelAttr);

                    // Include zone if this misc equipment is on and uses natural gas
                    if (equipmentOn == 1 && fuel.Equals("2-Natural gas"))
                    {
                        string zoneIdfName = zone.GetAttribute(IdfNameAttr);
                        string gasMeterName = "InteriorEquipment:NaturalGas:Zone:" + zoneIdfName;
                        meter.AddMeter(gasMeterName);
                        break; // One gas misc equipment flag is enough to include the zone
                    }
                }
            }

            return meter;
        }

        // Build a custom meter that sums electric equipment loads for the given zones
        private CustomMeter GetEquipmentElectricityMeter(string name, List<Zone> zones)
        {
            CustomMeter meter = new CustomMeter(name);

            foreach (Zone zone in zones)
            {
                bool electricEquipmentFound = false;

                // Check pure-electric equipment flags first
                for (int i = 0; i < PureElectricEquipmentAttrs.Length; i++)
                {
                    string onAttr = PureElectricEquipmentAttrs[i];
                    int equipmentOn = zone.GetAttributeAsInt(onAttr);

                    if (equipmentOn == 1)
                    {
                        electricEquipmentFound = true;
                        break;
                    }
                }

                // Check misc equipment that uses grid electricity
                for (int i = 0; i < MiscEquipmentAttrs.GetLength(0); i++) // rows
                {
                    string onAttr = MiscEquipmentAttrs[i, 0];
                    string fuelAttr = MiscEquipmentAttrs[i, 1];

                    int equipmentOn = zone.GetAttributeAsInt(onAttr);
                    string fuel = zone.GetAttribute(fuelAttr);
                    if (equipmentOn == 1 && fuel.Equals("1-Electricity from grid"))
                    {
                        electricEquipmentFound = true;
                        break;
                    }
                }

                // If any electric equipment is present, add the zone meter
                if (electricEquipmentFound)
                {
                    string zoneIdfName = zone.GetAttribute(IdfNameAttr);
                    string electricMeterName = "InteriorEquipment:Electricity:Zone:" + zoneIdfName;
                    meter.AddMeter(electricMeterName);
                }
            }

            return meter;
        }

        // Main routine: create all residential / non-residential meters and load them into IDF
        private void CreateCustomMeters(Building building)
        {
            List<Zone> residentialZones;
            List<Zone> nonResidentialZones;

            // Split zones into residential / non-residential
            GetResidentialAndNonResidentialZones(
                building,
                out residentialZones,
                out nonResidentialZones);

            // Lighting meters
            CustomMeter residentialLightsMeter = GetLightingMeter("ResidentialLights", residentialZones);
            CustomMeter nonResidentialLightsMeter = GetLightingMeter("Non-ResidentialLights", nonResidentialZones);

            // Gas equipment meters
            CustomMeter residentialGasEquipmentMeter = GetEquipmentGasMeter("ResidentialGas", residentialZones);
            CustomMeter nonResidentialGasEquipmentMeter = GetEquipmentGasMeter("Non-ResidentialGas", nonResidentialZones);

            // Electric equipment meters
            CustomMeter residentialElectricEquipmentMeter = GetEquipmentElectricityMeter("ResidentialElectricEquipment", residentialZones);
            CustomMeter nonResidentialElectricEquipmentMeter = GetEquipmentElectricityMeter("Non-ResidentialElectricEquipment", nonResidentialZones);

            // Collect all meters to process uniformly
            CustomMeter[] customMeters = new CustomMeter[]
            {
                residentialLightsMeter,
                nonResidentialLightsMeter,
                residentialGasEquipmentMeter,
                nonResidentialGasEquipmentMeter,
                residentialElectricEquipmentMeter,
                nonResidentialElectricEquipmentMeter
            };

            // Add non-empty custom meters to the IDF
            foreach (CustomMeter meter in customMeters)
            {
                if (!meter.IsEmpty())
                {
                    idfReader.Load(meter.ToIdfString());
                }
            }
        }
    }
}
