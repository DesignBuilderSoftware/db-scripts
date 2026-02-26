/*
Replace HeatExchanger:AirToAir:SensibleAndLatent with HeatExchanger:AirToAir:FlatPlate.
Object name needs to reference the HeatExchanger:AirToAir:SensibleAndLatent object name.
Attributes can be set in 'boilerplate' IDF text below.
*/

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
    string[] objectNames = new string[] { "Air Loop AHU Heat Recovery Device" };
    string hxBoilerPlate = @"
HeatExchanger:AirToAir:FlatPlate,
  {0},                        !- Name
  {1},                        !- Availability Schedule Name
  CounterFlow,                !- Flow Arrangement Type
  {2},                        !- Economizer Lockout
  1,                          !- Ratio of Supply to Secondary hA Values
  {3},                        !- Nominal Supply Air Flow Rate m3/s
  5.0,                        !- Nominal Supply Air Inlet Temperature C
  15.0,                       !- Nominal Supply Air Outlet Temperature C
  {3},                        !- Nominal Secondary Air Flow Rate m3/s
  20.0,                       !- Nominal Secondary Air Inlet Temperature C
  {4},                        !- Nominal Electric Power W
  {5},                        !- Supply Air Inlet Node Name
  {6},                        !- Supply Air Outlet Node Name
  {7},                        !- Secondary Air Inlet Node Name
  {8};                        !- Secondary Air Outlet Node Name";


    private IdfObject FindObject(IdfReader idfReader, string objectType, string objectName)
    {
      return idfReader[objectType].First(o => o[0] == objectName);
    }

    private void ReplaceObjectTypeInList(IdfReader idfReader, string listName, string oldObjectType, string oldObjectName, string newObjectType, string newObjectName)
    {
      IEnumerable<IdfObject> allEquipment = idfReader[listName];

      bool objectFound = false;

      foreach (IdfObject equipment in allEquipment)
      {
        if (!objectFound)
        {
          for (int i = 0; i < (equipment.Count - 1); i++)
          {
            Field field = equipment[i];
            Field nextField = equipment[i + 1];

            if (field.Value == oldObjectType && nextField.Value == oldObjectName)
            {
              field.Value = newObjectType;
              nextField.Value = newObjectName;
              objectFound = true;
              break;
            }
          }
        }
      }
    }

    public override void BeforeEnergySimulation()
    {
      IdfReader idfReader = new IdfReader(
        ApiEnvironment.EnergyPlusInputIdfPath,
        ApiEnvironment.EnergyPlusInputIddPath
        );

      string oldObjectType = "HeatExchanger:AirToAir:SensibleAndLatent";
      string newObjectType = "HeatExchanger:AirToAir:FlatPlate";

      foreach (string objectName in objectNames)
      {
        IdfObject hx = FindObject(idfReader, oldObjectType, objectName);

        ReplaceObjectTypeInList(idfReader, "AirLoopHVAC:OutdoorAirSystem:EquipmentList", oldObjectType, objectName, newObjectType, objectName);

        string name = hx["Name"].Value;
        string availability = hx["Availability Schedule Name"].Value;
        string lockout = hx["Economizer Lockout"].Value;
        string flowRate = hx["Nominal Supply Air Flow Rate"].Value;
        string electricPower = hx["Nominal Electric Power"].Value;
        string supplyInletNode = hx["Supply Air Inlet Node Name"].Value;
        string supplyOutletNode = hx["Supply Air Outlet Node Name"].Value;
        string extractInletNode = hx["Exhaust Air Inlet Node Name"].Value;
        string extractOutletNode = hx["Exhaust Air Outlet Node Name"].Value;

        string flatPlate = String.Format(hxBoilerPlate, name, availability, lockout, flowRate, electricPower, supplyInletNode, supplyOutletNode, extractInletNode, extractOutletNode);

        idfReader.Remove(hx);
        idfReader.Load(flatPlate);
      }
      idfReader.Save();
    }
  }
}