/*
Replace loop Demand Side Inlet(Outlet) Branch "Pipe:Adiabatic" component with "Pipe:Outdoor".

Material specification can be adjusted in the "pipeSpecs" variable content.

*/
using System.Runtime;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DB.Extensibility.Contracts;
using System;
using EpNet;

namespace DB.Extensibility.Scripts

{
  public class IdfFindAndReplace : ScriptBase, IScript
  {
    string pipeObject = @"  Pipe:Outdoor,
    {0},                     !- Name
    Insulated Pipe,          !- Construction Name
    {1},                     !- Fluid Inlet Node Name
    {2},                     !- Fluid Outlet Node Name
    Pipe OA Node,            !- Ambient Temperature Outdoor Air Node Name
    {3},                     !- Pipe Inside Diameter m
    {4};                     !- Pipe Length m";

   string pipeSpecs= @"  OutdoorAir:Node,
     Pipe OA Node,            !- Name
     -1.0;                    !- Height Above Ground m

 Construction,
    Insulated Pipe,          !- Name
    Pipe Insulation,         !- Outside Layer
    Pipe Steel;              !- Layer 2

  Material,
    Pipe Insulation,         !- Name
    VeryRough,               !- Roughness
    3.0E-02,                 !- Thickness m
    4.0E-02,                 !- Conductivity W/m-K
    91.0,                    !- Density kg/m3
    836.0,                   !- Specific Heat J/kg-K
    0.9,                     !- Thermal Absorptance
    0.5,                     !- Solar Absorptance
    0.5;                     !- Visible Absorptance

  Material,
    Pipe Steel,              !- Name
    Smooth,                  !- Roughness
    3.00E-03,                !- Thickness m
    45.31,                   !- Conductivity W/m-K
    7833.0,                  !- Density kg/m3
    500.0,                   !- Specific Heat J/kg-K
    0.9,                     !- Thermal Absorptance
    0.5,                     !- Solar Absorptance
    0.5;                     !- Visible Absorptance";

        enum Position
        {
            Outlet,
            Inlet
        }

       public IdfObject FindObject(IdfReader reader, string objectType, string objectName)
       {
           try
           {
               return reader[objectType].First(c => c[0] == objectName);
           }
           catch(Exception e)
           {
               throw new Exception(String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
           }
       }

        private void AddOutdoorPipe(IdfReader reader, string loopName, Position position, double pipeLength, double pipeDiameter)
        {
            string branchName;

            if (position == Position.Inlet) {
                branchName =  loopName + " Demand Side Inlet Branch";
            } else {
                branchName =  loopName + " Demand Side Outlet Branch";
            }

            IdfObject branch = FindObject(reader, "Branch", branchName);
            branch[2].Value = "Pipe:Outdoor";

            IdfObject adiabaticPipe = FindObject(reader, "Pipe:Adiabatic", branch[3].Value);
            string pipeName = adiabaticPipe[0].Value;
            string inletNode = adiabaticPipe[1].Value;
            string outletNode = adiabaticPipe[2].Value;

            string outdoorPipe = String.Format(pipeObject, pipeName, inletNode, outletNode, pipeDiameter, pipeLength);

            reader.Load(outdoorPipe);
            reader.Remove(adiabaticPipe);
        }


        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            AddOutdoorPipe(
                idfReader,
                loopName: "HW Loop",
                position: Position.Inlet,
                pipeLength: 100.0,
                pipeDiameter: 0.05
            );

            AddOutdoorPipe(idfReader, "HW Loop", Position.Outlet, 100.0, 0.05);

            idfReader.Load(pipeSpecs);
            idfReader.Save();
        }
    }
}