/*
    Description:
    Adds a SolarCollector:UnglazedTranspired object to the EnergyPlus input file.
    Integrates the solar collector into specified air handling units (AHUs) by updating the outdoor air system equipment list and building surface boundary conditions.
    Automatically configures output variables for solar collector performance reporting.

    Workflow (BeforeEnergySimulation):
    1. Initializes the IDF reader with EnergyPlus input and data dictionary files.
    2. Adds a free heating setpoint schedule (user can specify schedule name and setpoint value).
    3. Creates a SolarCollectorProperties object with user-defined parameters:
       - Collector name
       - Availability schedule name
       - Free heating setpoint schedule name
       - List of building surface names
       - (Optional physical properties can be customized)
    4. Defines air loop to control zone mappings via AirloopControlZone array (user specifies air loop and zone names).
    5. Integrates the solar collector into each air loop:
       - Updates the outdoor air system equipment list to include the collector.
       - Updates building surface boundary conditions to reference the collector's conditions model.
    6. Adds output variables for solar collector performance reporting.
    7. Saves all changes to the EnergyPlus input file.

    Expected User Inputs:
    - Solar collector name
    - Availability schedule name
    - Free heating setpoint schedule name
    - List of building surface names
    - Air loop and control zone names
    - (Optional) Physical properties for the collector

    DesignBuilder ssage:
    - Create a new C# script
    - Copy the content to the editor
    - Update air loop + control zone pairs and surface references
    - Enable the script and run a simulation
*/
using System.Runtime;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Globalization;
using DB.Extensibility.Contracts;
using System;
using EpNet;

namespace DB.Extensibility.Scripts

{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        private IdfReader Reader;

        public override void BeforeEnergySimulation()
        {
            Reader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            string freeHeatingScheduleName = "Free Heating Setpoint Schedule";
            AddFreeHeatingSetpointSchedule(freeHeatingScheduleName, 20.0);

            SolarCollectorProperties collectorProperties = new SolarCollectorProperties(
                "Unglazed Transpired Collector 1",
                "On 24/7",
                freeHeatingScheduleName,
                new List<string> { "BLOCK1:ZONE1_Wall_5_0_0", "BLOCK2:ZONE1_Wall_5_0_0" });

            AirloopControlZone[] airLoopControlZones = new AirloopControlZone[]
            {
                new AirloopControlZone { AirLoop = "Air Loop", ControlZone = "Block1:Zone1" },
                new AirloopControlZone { AirLoop = "Air Loop 1", ControlZone = "Block2:Zone1" }
            };
            AddUnglazedTranspiredMultisystem(collectorProperties, airLoopControlZones);
            AddSolarCollectorOutputs();

            Reader.Save();
        }

        public struct AirloopControlZone
        {
            public string AirLoop { get; set; }
            public string ControlZone { get; set; }
        }

        public struct MultisystemNodeSpecification
        {
            public string InletNode { get; set; }
            public string OutletNode { get; set; }
            public string MixedNode { get; set; }
            public string ZoneNode { get; set; }
        }

        public class SolarCollectorProperties
        {
            public string Name { get; set; }
            public string AvailabilityScheduleName { get; set; }
            public string FreeHeatingSetpointScheduleName { get; set; }

            public double DiameterOfPerforations { get; set; }
            public double DistanceBetweenPerforations { get; set; }
            public double ThermalEmissivity { get; set; }
            public double SolarAbsorptivity { get; set; }

            public double EffectiveOverallHeight { get; set; }
            public double EffectivePlenumGapThickness { get; set; }
            public double EffectivePlenumCrossSectionArea { get; set; }

            string HoleLayoutPattern { get; set; }
            string HeatExchangerEffectivenessCorrelation { get; set; }

            public double RatioOfActualCollectorAreaToGrossArea { get; set; }
            public string Roughness { get; set; }
            public double Thickness { get; set; }

            public double WindPerforationEffectiveness { get; set; }
            public double DischargeCoefficient { get; set; }
            public List<string> SurfaceNames { get; set; }

            public SolarCollectorProperties(string name,
                string availabilityScheduleName,
                string setpointScheduleName,
                List<string> surfaceNames,

                double perforationDiameter = 0.003,
                double perforationSpacing = 0.05,
                double thermalEmissivity = 0.9,
                double solarAbsorptivity = 0.85,

                double effectiveOverallHeight = 4.0,
                double effectivePlenumGapThickness = 0.1,
                double effectivePlenumCrossSectionArea = 5.0,

                string holeLayoutPattern = "Triangle",
                string heatExchangeEffectivenessCorrelation = "Kutscher1994",

                double ratioOfActualCollectorAreaToGrossArea = 1.165,
                string collectorRoughness = "MediumRough",
                double collectorThickness = 0.001,
                double windPerforationEffectiveness = 0.25,
                double dischargeCoefficient = 0.65)
            {
                Name = name;
                AvailabilityScheduleName = availabilityScheduleName;
                FreeHeatingSetpointScheduleName = setpointScheduleName;
                SurfaceNames = surfaceNames;

                DiameterOfPerforations = perforationDiameter;
                DistanceBetweenPerforations = perforationSpacing;
                ThermalEmissivity = thermalEmissivity;
                SolarAbsorptivity = solarAbsorptivity;

                EffectiveOverallHeight = effectiveOverallHeight;
                EffectivePlenumGapThickness = effectivePlenumGapThickness;
                EffectivePlenumCrossSectionArea = effectivePlenumCrossSectionArea;

                HoleLayoutPattern = holeLayoutPattern;
                HeatExchangerEffectivenessCorrelation = heatExchangeEffectivenessCorrelation;

                RatioOfActualCollectorAreaToGrossArea = ratioOfActualCollectorAreaToGrossArea;
                Roughness = collectorRoughness;
                Thickness = collectorThickness;

                WindPerforationEffectiveness = windPerforationEffectiveness;
                DischargeCoefficient = dischargeCoefficient;
            }

            public string GetIdfString()
            {
                string solarCollectorTemplate = @"  SurfaceProperty:OtherSideConditionsModel,
    {0} Conditions Model,    !- Name
    GapConvectionRadiation;  !- Type of Modeling

  SolarCollector:UnglazedTranspired,
    {0},                     !- Name
    {0} Conditions Model,    !- Boundary Conditions Model Name
    {1},                     !- Availability Schedule Name
    ,                        !- Inlet Node Name
    ,                        !- Outlet Node Name
    ,                        !- Setpoint Node Name
    ,                        !- Zone Node Name
    {2},                     !- Free Heating Setpoint Schedule Name
    {3},                     !- Diameter of Perforations in Collector [m]
    {4},                     !- Distance Between Perforations in Collector [m]
    {5},                     !- Thermal Emissivity of Collector Surface [dimensionless]
    {6},                     !- Solar Absorbtivity of Collector Surface [dimensionless]
    {7},                     !- Effective Overall Height of Collector
    {8},                     !- Effective Gap Thickness of Plenum Behind Collector [m]
    {9},                     !- Effective Cross Section Area of Plenum Behind Collector [m2]
    {10},                    !- Hole Layout Pattern for Pitch [square, triangle]
    {11},                    !- Heat Exchange Effectiveness Correlation [VanDeckerHollandsBrunger2001, Kutscher1994]
    {12},                    !- Ratio of Actual Collector Surface Area to Projected Surface Area [dimensionless]
    {13},                    !- Roughness of Collector
    {14},                    !- Collector Thickness [m]
    {15},                    !- Effectiveness for Perforations with Respect to Wind [dimensionless]
    {16},                    !- Discharge Coefficient for Openings with Respect to Buoyancy Driven Flow [dimensionless]";

                for (int i = 0; i < SurfaceNames.Count; i++)
                {
                    bool isLast = (i == SurfaceNames.Count - 1);
                    string terminator = isLast ? ";" : ",";
                    solarCollectorTemplate += string.Format(CultureInfo.InvariantCulture,
                        ",\n    {0}{1}                     !- Surface {2} Name",
                        SurfaceNames[i],
                        terminator,
                        i + 1);
                }

                return string.Format(CultureInfo.InvariantCulture, solarCollectorTemplate,
                    Name,                                                                    // {0}
                    AvailabilityScheduleName,                                               // {1}
                    FreeHeatingSetpointScheduleName,                                        // {2}
                    DiameterOfPerforations.ToString("F3", CultureInfo.InvariantCulture),   // {3}
                    DistanceBetweenPerforations.ToString("F3", CultureInfo.InvariantCulture), // {4}
                    ThermalEmissivity.ToString("F3", CultureInfo.InvariantCulture),        // {5}
                    SolarAbsorptivity.ToString("F3", CultureInfo.InvariantCulture),        // {6}
                    EffectiveOverallHeight.ToString("F3", CultureInfo.InvariantCulture),   // {7}
                    EffectivePlenumGapThickness.ToString("F3", CultureInfo.InvariantCulture), // {8}
                    EffectivePlenumCrossSectionArea.ToString("F3", CultureInfo.InvariantCulture), // {9}
                    HoleLayoutPattern,                                                      // {10}
                    HeatExchangerEffectivenessCorrelation,                                 // {11}
                    RatioOfActualCollectorAreaToGrossArea.ToString("F3", CultureInfo.InvariantCulture), // {12}
                    Roughness,                                                             // {13}
                    Thickness.ToString("F3", CultureInfo.InvariantCulture),               // {14}
                    WindPerforationEffectiveness.ToString("F3", CultureInfo.InvariantCulture), // {15}
                    DischargeCoefficient.ToString("F3", CultureInfo.InvariantCulture)     // {16}
                );
            }
        }
        
        public class SolarCollectorMultisystem
        {
            public string SolarCollectorName { get; set; }
            public List<MultisystemNodeSpecification> NodeSpecification = new List<MultisystemNodeSpecification>();

            public SolarCollectorMultisystem(string solarCollectorName)
            {
                SolarCollectorName = solarCollectorName;
            }

            public string GetIdfString()
            {
                string template = @"  SolarCollector:UnglazedTranspired:Multisystem,
    {0},                    !- Solar Collector Name" + Environment.NewLine;

                string nodesTemplate = @"    {0},               !- Outdoor Air System {5} Collector Inlet Node
    {1},                !- Outdoor Air System {5} Collector Outlet Node
    {2},                !- Outdoor Air System {5} Mixed Air Node
    {3}{4}              !- Outdoor Air System {5} Zone Node" + Environment.NewLine;

                for (int i = 0; i < NodeSpecification.Count; i++)
                {
                    bool isLast = (i == NodeSpecification.Count - 1);
                    string terminator = isLast ? ";" : ",";
                    MultisystemNodeSpecification specification = NodeSpecification[i];
                    template += string.Format(nodesTemplate,
                        specification.InletNode,
                        specification.OutletNode,
                        specification.MixedNode,
                        specification.ZoneNode,
                        terminator,
                        i + 1);
                }

                return string.Format(template, SolarCollectorName);
            }
        }

        public void AddFreeHeatingSetpointSchedule(string name, double setpoint)
        {
            string template = @"  Schedule:Compact,
    {0},                     !- Name
    Temperature,             !- Schedule Type Limits Name
    Through: 12/31,          !- Field 1
    For: AllDays,            !- Field 2
    Until: 24:00,{1};        !- Field 3";
            Reader.Load(string.Format(CultureInfo.InvariantCulture, template, name, setpoint));
        }

        public void AddUnglazedTranspiredMultisystem(SolarCollectorProperties collectorProperties, AirloopControlZone[] airloopControlZones)
        {
            SolarCollectorMultisystem multisystem = new SolarCollectorMultisystem(collectorProperties.Name);
            foreach (var airloopControlZone in airloopControlZones)
            {
                string airLoopName = airloopControlZone.AirLoop;
                string controlZoneNode = airloopControlZone.ControlZone;

                MultisystemNodeSpecification nodeSpec = AddTranspiredCollectorToAirLoop(collectorProperties.Name, airLoopName, controlZoneNode);
                multisystem.NodeSpecification.Add(nodeSpec);
            }

            UpdateBuildingSurfaceConditions(collectorProperties.SurfaceNames, collectorProperties.Name);

            Reader.Load(collectorProperties.GetIdfString());
            Reader.Load(multisystem.GetIdfString());
        }

        public void UpdateBuildingSurfaceConditions(List<String> surfaceNames, string collectorName)
        {
            string conditionsModelName = collectorName + " Conditions Model";
            foreach (string surfaceName in surfaceNames)
            {
                IdfObject surface = FindObject("BuildingSurface:Detailed", surfaceName);
                surface["Outside Boundary Condition"].Value = "OtherSideConditionsModel";
                surface["Outside Boundary Condition Object"].Value = conditionsModelName;
            }
        }

        public MultisystemNodeSpecification AddTranspiredCollectorToAirLoop(string solarCollectorName, string airLoopName, string controlZoneName)
        {
            string oaControllerName = airLoopName + " AHU Outdoor Air Controller";
            IdfObject oaController = FindObject("Controller:OutdoorAir", oaControllerName);

            string inletNode = oaController["Actuator Node Name"].Value;
            string mixedNode = oaController["Mixed Air Node Name"].Value;
            string outletNode = airLoopName + " UTSC Outlet Node";
            string zoneNode = controlZoneName + " Zone Air Node";

            MultisystemNodeSpecification nodeSpec = new MultisystemNodeSpecification
            {
                InletNode = inletNode,
                OutletNode = outletNode,
                MixedNode = mixedNode,
                ZoneNode = zoneNode
            };

            string oaEquipmentName = airLoopName + " AHU Outdoor air Equipment List";

            IdfObject oaEquipmentList = FindObject("AirLoopHVAC:OutdoorAirSystem:EquipmentList", oaEquipmentName);

            string firstComponentType = oaEquipmentList["Component 1 Object Type"].Value;
            string firstComponentName = oaEquipmentList["Component 1 Name"].Value;
            IdfObject firstComponent = FindObject(firstComponentType, firstComponentName);

            if (firstComponentType.Equals("HeatExchanger:AirToAir:SensibleAndLatent", StringComparison.OrdinalIgnoreCase))
            {
                firstComponent["Supply Air Inlet Node Name"].Value = outletNode;
            }
            else if (firstComponentType.Equals("OutdoorAir:Mixer", StringComparison.OrdinalIgnoreCase))
            {
                firstComponent["Outdoor Air Stream Node Name"].Value = outletNode;
            }
            else
            {
                throw new Exception("Unexpected component in the Outdoor Air Equipment: " + firstComponentName);
            }

            oaEquipmentList.InsertFields(1, "SolarCollector:UnglazedTranspired", solarCollectorName);

            return nodeSpec;
        }

        public void AddSolarCollectorOutputs()
        {
            string variables = @"Output:Variable,*,Solar Collector Heat Exchanger Effectiveness,hourly; !- HVAC Average []
Output:Variable,*,Solar Collector Leaving Air Temperature,hourly; !- HVAC Average [C]
Output:Variable,*,Solar Collector Outside Face Suction Velocity,hourly; !- HVAC Average [m/s]
Output:Variable,*,Solar Collector Surface Temperature,hourly; !- HVAC Average [C]
Output:Variable,*,Solar Collector Plenum Air Temperature,hourly; !- HVAC Average [C]
Output:Variable,*,Solar Collector Sensible Heating Rate,hourly; !- HVAC Average [W]
Output:Variable,*,Solar Collector Sensible Heating Energy,hourly; !- HVAC Sum [J]
Output:Variable,*,Solar Collector Natural Ventilation Air Change Rate,hourly; !- HVAC Average [ach]
Output:Variable,*,Solar Collector Natural Ventilation Mass Flow Rate,hourly; !- HVAC Average [kg/s]
Output:Variable,*,Solar Collector Wind Natural Ventilation Mass Flow Rate,hourly; !- HVAC Average [kg/s]
Output:Variable,*,Solar Collector Buoyancy Natural Ventilation Mass Flow Rate,hourly; !- HVAC Average [kg/s]
Output:Variable,*,Solar Collector Incident Solar Radiation,hourly; !- HVAC Average [W/m2]
Output:Variable,*,Solar Collector System Efficiency,hourly; !- HVAC Average []
Output:Variable,*,Solar Collector Surface Efficiency,hourly; !- HVAC Average []";
            Reader.Load(variables);
        }

        public IdfObject FindObject(string objectType, string objectName)
        {
            try
            {
                return Reader[objectType].First(c => c[0] == objectName);
            }
            catch (Exception)
            {
                throw new Exception(String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }
    }
}