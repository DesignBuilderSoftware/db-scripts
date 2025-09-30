/*
Replace electric heating coil in all generic air handling units with "desuperheater" coils.

The AHU needs to contain DX cooling coil and use "desuperheater" string in its name.

*/

using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System;
using DB.Extensibility.Contracts;
using EpNet;


namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {

        private IdfObject FindObject(IdfReader reader, string objectType, string objectName) {
             return reader[objectType].First(c => c[0] == objectName);
        }  

        private void UpdateBranches(IdfReader idfReader)  
        {
             IEnumerable<IdfObject> branches = idfReader["Branch"];
             foreach (IdfObject branch in branches) 
             {
                 if (branch[0].Value.Contains("desuperheater") && branch[0].Value.Contains("AHU Main Branch"))
                 {
                      UpdateBranch(idfReader, branch);                   
                 }
             }
        }

        private void UpdateBranch(IdfReader reader, IdfObject branch)
        {
             const string DesuperHeaterObject = "Coil:Heating:Desuperheater";
             const string ElectricCoilObject = "Coil:Heating:Electric";
             const string DxSystemObjectName = "CoilSystem:Cooling:DX";
             IdfObject DxCoil = new IdfObject();
             bool DxIncluded = false;

             foreach (int i in Enumerable.Range(1, (branch.Count - 2)))
             {     
                   Field ThisField = branch.Fields[i];
                   Field NextField = branch.Fields[i + 1];
                   if (ThisField.Value.Contains(DxSystemObjectName))
                   { 
                        MessageBox.Show("DX Coil System included " + branch[0]);
                        DxIncluded = true;
                        IdfObject CoilSystem = FindObject(reader, DxSystemObjectName, NextField.Value);
                        DxCoil = FindObject(reader, CoilSystem["Cooling Coil Object Type"], CoilSystem["Cooling Coil Name"]);
                        
                   }

                   if (DxIncluded && ThisField.Comment.ToLower().Contains("object type") && ThisField.Value.ToLower() == ElectricCoilObject.ToLower())
                   {
                        IdfObject ElectricCoil = FindObject(reader, ElectricCoilObject, NextField.Value);
                        ThisField.Value = DesuperHeaterObject;
                        AddDesuperheater(reader, ElectricCoil, DxCoil);
                        MessageBox.Show("Replacing electric heating coil: " + ElectricCoil[0] + " with desuperheater coil." );
                        reader.Remove(ElectricCoil);
                        break;
                   }
             } 
        }
        private void AddDesuperheater(IdfReader reader, IdfObject electricCoil, IdfObject dxCoil)
        {
            string name = electricCoil[0].Value;
            string availability = electricCoil[1].Value;
            string inlet = electricCoil[4].Value;
            string outlet = electricCoil[5].Value;
            string dxClass = dxCoil.IdfClass;
            string dxName = dxCoil[0].Value;
            string tempNode = electricCoil[6].Value;
            double efficiency = 0.3;          
             
            string desuperheater = String.Format("Coil:Heating:Desuperheater,{0},{1},{2},{3},{4},{5},{6},{7},0;", name, availability, efficiency, inlet, outlet, dxClass, dxName, tempNode);

           reader.Load(desuperheater);
        }  
         public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);
            
            UpdateBranches(idfReader);

            idfReader.Save();
        }
    }
}
