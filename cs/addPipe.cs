/*
Add indoor and outdoor pipe objects

Purpose:
This DesignBuilder C# script adds Pipe:Indoor and Pipe:Outdoor Objects to the EnergyPlus IDF

Main Steps:
1) Create insulated pipe constructions (Construction + Material layers)
2) Insert Pipe:Indoor and Pipe:Outdoor objects into specified plant Branch objects
3) UpdatePlantLoop outlet node references to match the new pipe outlet nodes

How to Use:

Configuration
Open BeforeEnergySimulation() and edit the "Configuration" section:
- Pipe constructions:
  - pipeConstructionName: name for the Construction and associated Material layers
  - insulationThickness: insulation layer thickness [m]
- Indoor pipes:
  - zoneName: the thermal zone used for ambient temperature (Environment Type = ZONE)
- Indoor/Outdoor pipes:
  - branchName: the target Branch to modify
  - pipeName: name of the pipe object that will be inserted
  - pipeConstructionName: Construction name to assign
  - pipeInsideDiameter: inside diameter [m]
  - pipeLength: length [m]

Prerequisites / Placeholders
- The IDF must contain:
  - Branch objects with names matching branchName inputs.
  - PlantLoop objects (if you expect PlantLoop outlet node references to be updated).
  - For Pipe:Indoor, the specified zoneName must exist (Zone object).
- NOTE: The inlet node name for the inserted pipe is taken from the last field in the Branch object (see AddIndoorPipeToBranch / AddOutdoorPipeToBranch).

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. 
Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Linq;
using System.Windows.Forms;
using System.Globalization;

using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class AddPipes : ScriptBase, IScript
    {
        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // ----------------------------
            // USER CONFIGURATION SECTION
            // ----------------------------

            // Define pipe constructions (Construction + Material layers)
            string indoorPipeConstructionName = "pipe construction 1";
            AddInsulatedPipeConstruction(idfReader, indoorPipeConstructionName, insulationThickness: 0.025);

            string outdoorPipeConstructionName = "pipe construction 2";
            AddInsulatedPipeConstruction(idfReader, outdoorPipeConstructionName, insulationThickness: 0.05);

            // Define pipes to be added (target by Branch name)
            AddIndoorPipeToBranch(
                reader: idfReader,
                zoneName: "BASEMENT:ZONE1",
                branchName: "HW Loop Demand Side Inlet Branch",
                pipeName: "pipe 1",
                pipeConstructionName: indoorPipeConstructionName,
                pipeInsideDiameter: 0.03,
                pipeLength: 30);

            AddIndoorPipeToBranch(
                reader: idfReader,
                zoneName: "BASEMENT:ZONE1",
                branchName: "HW Loop Demand Side Outlet Branch",
                pipeName: "pipe 2",
                pipeConstructionName: indoorPipeConstructionName,
                pipeInsideDiameter: 0.03,
                pipeLength: 30);

            AddOutdoorPipeToBranch(
                reader: idfReader,
                branchName: "HW Loop Supply Side Inlet Branch",
                pipeName: "pipe 3",
                pipeConstructionName: outdoorPipeConstructionName,
                pipeInsideDiameter: 0.03,
                pipeLength: 50);

            AddOutdoorPipeToBranch(
                reader: idfReader,
                branchName: "HW Loop Supply Side Outlet Branch",
                pipeName: "pipe 4",
                pipeConstructionName: outdoorPipeConstructionName,
                pipeInsideDiameter: 0.03,
                pipeLength: 50);

            idfReader.Save();
        }

        private IdfObject GetRequiredObject(IdfReader reader, string objectType, string objectName)
        {
            try
            {
                return reader[objectType].First(c => c[0] == objectName);
            }
            catch (Exception)
            {
                throw new MissingFieldException(
                    string.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        // Updates PlantLoop outlet node name fields if they match a provided node name.
        private void UpdatePlantLoopOutletNodeReferences(IdfReader reader, string oldOutletNodeName, string newOutletNodeName)
        {
            string[] plantLoopOutletFieldNames = new string[]
            {
                "Plant Side Outlet Node Name",
                "Demand Side Outlet Node Name"
            };

            foreach (var plantLoop in reader["PlantLoop"])
            {
                foreach (var fieldName in plantLoopOutletFieldNames)
                {
                    if (string.Equals(plantLoop[fieldName].Value, oldOutletNodeName, StringComparison.OrdinalIgnoreCase))
                    {
                        plantLoop[fieldName].Value = newOutletNodeName;
                    }
                }
            }
        }

        private string BuildIndoorPipeIdfText(
            string pipeName,
            string inletNodeName,
            string outletNodeName,
            string pipeConstructionName,
            string zoneName,
            double pipeInsideDiameter,
            double pipeLength)
        {
            string pipeTemplate = @"  
Pipe:Indoor,
    {0},                     !- Name
    {1},                     !- Construction Name
    {2},                     !- Fluid Inlet Node Name
    {3},                     !- Fluid Outlet Node Name
    ZONE,                    !- Environment Type
    {4},                     !- Ambient Temperature Zone Name
    ,                        !- Ambient Temperature Schedule Name
    ,                        !- Ambient Air Velocity Schedule Name
    {5},                     !- Pipe Inside Diameter m
    {6};                     !- Pipe Length m";

            return string.Format(
                CultureInfo.InvariantCulture,
                pipeTemplate,
                pipeName,
                pipeConstructionName,
                inletNodeName,
                outletNodeName,
                zoneName,
                pipeInsideDiameter,
                pipeLength);
        }

        private string BuildOutdoorPipeIdfText(
            string pipeName,
            string inletNodeName,
            string outletNodeName,
            string pipeConstructionName,
            double pipeInsideDiameter,
            double pipeLength)
        {
            string pipeTemplate = @"  
Pipe:Outdoor,
    {0},                    !- Construction Name
    {1},                    !- Fluid Inlet Node Name
    {2},                    !- Fluid Outlet Node Name
    {3},                    !- Comp1 Outlet Node Name
    {0} Outdoor Air Node,   !- Ambient Temperature Outdoor Air Node Name
    {4},                    !- Pipe Inside Diameter
    {5};                    !- pipe length

OutdoorAir:Node,
    {0} Outdoor Air Node;   !- Name";

            return string.Format(
                CultureInfo.InvariantCulture,
                pipeTemplate,
                pipeName,
                pipeConstructionName,
                inletNodeName,
                outletNodeName,
                pipeInsideDiameter,
                pipeLength);
        }

        private void AddIndoorPipeToBranch(
            IdfReader reader,
            string zoneName,
            string branchName,
            string pipeName,
            string pipeConstructionName,
            double pipeInsideDiameter,
            double pipeLength)
        {
            try
            {
                IdfObject branch = GetRequiredObject(reader, "Branch", branchName);

                // NOTE: This assumes the last field in the Branch is the node name to use as the pipe inlet node.
                string inletNodeName = branch[branch.Count - 1].Value;
                string newOutletNodeName = pipeName + " Outlet Node";

                string pipeIdfText = BuildIndoorPipeIdfText(
                    pipeName,
                    inletNodeName,
                    newOutletNodeName,
                    pipeConstructionName,
                    zoneName,
                    pipeInsideDiameter,
                    pipeLength);

                UpdatePlantLoopOutletNodeReferences(reader, inletNodeName, newOutletNodeName);

                branch.AddFields("Pipe:Indoor", pipeName, inletNodeName, newOutletNodeName);
                reader.Load(pipeIdfText);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private void AddOutdoorPipeToBranch(
            IdfReader reader,
            string branchName,
            string pipeName,
            string pipeConstructionName,
            double pipeInsideDiameter,
            double pipeLength)
        {
            try
            {
                IdfObject branch = GetRequiredObject(reader, "Branch", branchName);

                // NOTE: This assumes the last field in the Branch is the node name to use as the pipe inlet node.
                string inletNodeName = branch[branch.Count - 1].Value;

                string newOutletNodeName = pipeName + " Outlet Node";

                string pipeIdfText = BuildOutdoorPipeIdfText(
                    pipeName,
                    inletNodeName,
                    newOutletNodeName,
                    pipeConstructionName,
                    pipeInsideDiameter,
                    pipeLength);

                // Redirect PlantLoop outlet node references if they match the old node
                UpdatePlantLoopOutletNodeReferences(reader, inletNodeName, newOutletNodeName);

                branch.AddFields("Pipe:Outdoor", pipeName, inletNodeName, newOutletNodeName);
                reader.Load(pipeIdfText);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private void AddInsulatedPipeConstruction(IdfReader reader, string pipeConstructionName, double insulationThickness)
        {
            // Construction is composed of an insulation layer + steel layer
            string pipeInsulationMaterialName = pipeConstructionName + " insulation";
            string pipeSteelMaterialName = pipeConstructionName + " steel";

            string pipeConstructionTemplate = @"  
Construction,
    {0},                     !-Name
    {1},                     !-Outside Layer
    {2};                     !-Layer 2

Material,
    {1},                     !- Name
    VeryRough,               !- Roughness
    {3},                     !- Thickness m
    4.0E-02,                 !- Conductivity W/m-K
    91.0,                    !- Density kg/m3
    836.0,                   !- Specific Heat J/kg-K
    0.9,                     !- Thermal Absorptance
    0.5,                     !- Solar Absorptance
    0.5;                     !- Visible Absorptance

Material,
    {2},                     !- Name
    Smooth,                  !- Roughness
    3.00E-03,                !- Thickness m
    45.31,                   !- Conductivity W/m-K
    7833.0,                  !- Density kg/m3
    500.0,                   !- Specific Heat J/kg-K
    0.9,                     !- Thermal Absorptance
    0.5,                     !- Solar Absorptance
    0.5;                     !- Visible Absorptance
";

            string pipeConstructionIdfText = string.Format(
                CultureInfo.InvariantCulture,
                pipeConstructionTemplate,
                pipeConstructionName,
                pipeInsulationMaterialName,
                pipeSteelMaterialName,
                insulationThickness);

            reader.Load(pipeConstructionIdfText);
        }
    }
}