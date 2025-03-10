/*
This C# script is designed to enhance the functionality by providing an option to add Pipe:Indoor and Pipe:Outdoor objects to an IDF file. 

*/
using System.Runtime;
using System;
using System.Linq;
using System.Windows.Forms;

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

            // define pump constructions
            string pipeConstruction1 = "pipe construction 1";
            AddInsulatedPipe(idfReader, pipeConstruction1, 0.025);

            string pipeConstruction2 = "pipe construction 2";
            AddInsulatedPipe(idfReader, pipeConstruction2, 0.05);

            // define pipes to be added
            AddIndoorPipe(
                reader: idfReader,
                zoneName: "BASEMENT:ZONE1",
                branchName: "HW Loop Demand Side Inlet Branch",
                pipeName: "pipe 1",
                pipeConstructionName: pipeConstruction1,
                pipeInsideDiameter: 0.03,
                pipeLength: 30);

            AddIndoorPipe(
                reader: idfReader,
                zoneName: "BASEMENT:ZONE1",
                branchName: "HW Loop Demand Side Outlet Branch",
                pipeName: "pipe 2",
                pipeConstructionName: pipeConstruction1,
                pipeInsideDiameter: 0.03,
                pipeLength: 30);

            AddOutdoorPipe(
                reader: idfReader,
                branchName: "HW Loop Supply Side Inlet Branch",
                pipeName: "pipe 3",
                pipeConstructionName: pipeConstruction2,
                pipeInsideDiameter: 0.03,
                pipeLength: 50);

            AddOutdoorPipe(
                reader: idfReader,
                branchName: "HW Loop Supply Side Outlet Branch",
                pipeName: "pipe 4",
                pipeConstructionName: pipeConstruction2,
                pipeInsideDiameter: 0.03,
                pipeLength: 50);


            idfReader.Save();
        }

        public IdfObject FindObject(IdfReader reader, string objectType, string objectName)
        {
            try
            {
                return reader[objectType].First(c => c[0] == objectName);
            }
            catch (Exception e)
            {
                throw new MissingFieldException(String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        public void UpdatePlantLoopNode(IdfReader reader, string originalOutletNodeName, string newNodeName)
        {
            string[] fieldNames = new string[] { "Plant Side Outlet Node Name", "Demand Side Outlet Node Name" };
            foreach (var plantLoop in reader["PlantLoop"])
            {
                foreach (var fieldName in fieldNames)
                {
                    if (string.Equals(plantLoop[fieldName].Value, originalOutletNodeName,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        plantLoop[fieldName].Value = newNodeName;
                    }
                }
            }
        }


        public string GetIndoorPipe(string name, string inletNodeName, string outletNodeName, string pipeConstructionName, string zoneName, double pipeInsideDiameter, double pipeLength)
        {
            string pipe = @"  Pipe:Indoor,
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
            return String.Format(pipe, name, pipeConstructionName, inletNodeName, outletNodeName, zoneName, pipeInsideDiameter, pipeLength);
        }

        public string GetOutdoorPipe(string name, string inletNodeName, string outletNodeName, string pipeConstructionName, double pipeInsideDiameter, double pipeLength)
        {
            string pipe = @"  Pipe:Outdoor,
    {0},                    !- Construction Name
    {1},                    !- Fluid Inlet Node Name
    {2},                    !- Fluid Outlet Node Name
    {3},                    !- Comp1 Outlet Node Name
    {0} Outdoor Air Node,   !- Ambient Temperature Outdoor Air Node Name
    {4},                    !- Pipe Inside Diameter
    {5};                    !- pipe length

  OutdoorAir:Node,
    {0} Outdoor Air Node;   !- Name";
            return String.Format(pipe, name, pipeConstructionName, inletNodeName, outletNodeName, pipeInsideDiameter, pipeLength);
        }


        public void AddIndoorPipe(IdfReader reader, string zoneName, string branchName, string pipeName, string pipeConstructionName, double pipeInsideDiameter, double pipeLength)
        {
            try
            {
                IdfObject branch = FindObject(reader, "Branch", branchName);
                string inletNodeName = branch[branch.Count - 1].Value;
                string outletNodeName = pipeName + " Outlet Node";
                string pipe = GetIndoorPipe(pipeName, inletNodeName, outletNodeName, pipeConstructionName, zoneName, pipeInsideDiameter, pipeLength);

                UpdatePlantLoopNode(reader, inletNodeName, outletNodeName);

                branch.AddFields("Pipe:Indoor", pipeName, inletNodeName, outletNodeName);
                reader.Load(pipe);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        public void AddOutdoorPipe(IdfReader reader, string branchName, string pipeName, string pipeConstructionName, double pipeInsideDiameter, double pipeLength)
        {
            try
            {
                IdfObject branch = FindObject(reader, "Branch", branchName);
                string inletNodeName = branch[branch.Count - 1].Value;
                string outletNodeName = pipeName + " Outlet Node";
                string pipe = GetOutdoorPipe(pipeName, inletNodeName, outletNodeName, pipeConstructionName, pipeInsideDiameter, pipeLength);

                UpdatePlantLoopNode(reader, inletNodeName, outletNodeName);

                branch.AddFields("Pipe:Outdoor", pipeName, inletNodeName, outletNodeName);
                reader.Load(pipe);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        public void AddInsulatedPipe(IdfReader reader, string pipeConstructionName, double insulationThickness)
        {
            string pipeInsulationName = pipeConstructionName + " insulation";
            string pipeSteelName = pipeConstructionName + " steel";
            string pipeConstructionTemplate = @"  Construction,
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
            string pipeConstruction = String.Format(pipeConstructionTemplate, pipeConstructionName, pipeInsulationName, pipeSteelName, insulationThickness);
            reader.Load(pipeConstruction);
        }
    }
}