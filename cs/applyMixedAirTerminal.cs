/*
CAV to Mixer Replacement for Zones with VRF Terminal Units

This DesignBuilder C# script modifies the EnergyPlus IDF by replacing AirTerminal:SingleDuct:ConstantVolume:NoReheat (CAV NoReheat) 
terminals with AirTerminal:SingleDuct:Mixer terminals, in zones that contain a ZoneHVAC:TerminalUnit:VariableRefrigerantFlow (VRF terminal unit).

Purpose (main steps)
The goal is to allow mixed air delivery at the zone inlet by combining:
- Primary air from a central system (e.g., DOAS) via the terminal inlet node
- Secondary air derived from the VRF terminal unit inlet node (treated here as the mixer secondary inlet)

1) Scan ZoneHVAC:EquipmentList objects and find zones that include both:
   - A ZoneHVAC:AirDistributionUnit referencing an AirTerminal:SingleDuct:ConstantVolume:NoReheat terminal
   - A ZoneHVAC:TerminalUnit:VariableRefrigerantFlow terminal unit
2) For each matching zone:
   - Create and insert an AirTerminal:SingleDuct:Mixer object
   - Update ZoneHVAC:AirDistributionUnit to reference the Mixer terminal type
   - Rewire the VRF cooling coil inlet node to the (old) terminal outlet node
   - Remove the old CAV NoReheat terminal and outlet node from the zone inlet NodeList (if present)

How to Use

Configuration
- Target objects are detected by object types (constants at top of the script):
  - Old terminal type: AirTerminal:SingleDuct:ConstantVolume:NoReheat
  - New terminal type: AirTerminal:SingleDuct:Mixer
  - Zone unit type: ZoneHVAC:TerminalUnit:VariableRefrigerantFlow
  - Distribution unit type: ZoneHVAC:AirDistributionUnit

Prerequisites (required placeholders)
Base model must contain, for each target zone:
- ZoneHVAC:EquipmentList that includes:
  - ZoneHVAC:AirDistributionUnit (points to an AirTerminal:SingleDuct:ConstantVolume:NoReheat)
  - ZoneHVAC:TerminalUnit:VariableRefrigerantFlow

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System;
using System.Collections.Generic;
using System.Linq;

using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        private IdfReader idf;

        // Object types used to identify targets and apply replacements
        private const string DistributionUnitType = "ZoneHVAC:AirDistributionUnit";
        private const string OldTerminalType = "AirTerminal:SingleDuct:ConstantVolume:NoReheat";
        private const string NewTerminalType = "AirTerminal:SingleDuct:Mixer";
        private const string VrfTerminalUnitType = "ZoneHVAC:TerminalUnit:VariableRefrigerantFlow";

        // IDF template for the replacement Mixer terminal
        private static readonly string MixerTerminalTemplate = @"
AirTerminal:SingleDuct:Mixer,
    {0},                   !- Name
    {1},                   !- ZoneHVAC Unit Object Type
    {2},                   !- ZoneHVAC Unit Object Name
    {3},                   !- Mixer Outlet Node Name
    {4},                   !- Mixer Primary Air Inlet Node Name
    {5},                   !- Mixer Secondary Air Inlet Node Name
    InletSide,             !- Mixer Connection Type
    {6},                   !- Design Specification Outdoor Air Object Name
    ;                      !- Per Person Ventilation Rate Mode";

        public override void BeforeEnergySimulation()
        {
            idf = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            ReplaceCavTerminalsWithMixers();

            idf.Save();
        }

        // Main routine: find target zones and replace the CAV terminal with a Mixer terminal
        private void ReplaceCavTerminalsWithMixers()
        {
            // Find zones that contain BOTH the old terminal (CAV:NoReheat) and a VRF terminal unit
            List<ZoneEquipmentSpecification> zoneSpecs = FindZonesWithCavAndVrf(OldTerminalType, VrfTerminalUnitType);

            foreach (ZoneEquipmentSpecification zoneSpec in zoneSpecs)
            {
                IdfObject oldTerminal = FindObject(zoneSpec.TerminalType, zoneSpec.TerminalName);
                IdfObject vrfTerminalUnit = FindObject(zoneSpec.ZoneUnitType, zoneSpec.ZoneUnitName);

                string secondaryInletNodeName = vrfTerminalUnit["Terminal Unit Air Inlet Node Name"].Value;

                // Build the new Mixer with required nodes
                string mixerIdfText = String.Format(
                    MixerTerminalTemplate,
                    oldTerminal["Name"].Value,
                    zoneSpec.ZoneUnitType,
                    zoneSpec.ZoneUnitName,
                    oldTerminal["Air Outlet Node Name"].Value,
                    oldTerminal["Air Inlet Node Name"].Value,
                    secondaryInletNodeName,
                    oldTerminal["Design Specification Outdoor Air Object Name"].Value);

                IdfObject distributionUnit = FindObject(DistributionUnitType, zoneSpec.DistributionUnitName);
                distributionUnit["Air Terminal Object Type"].Value = NewTerminalType;

                IdfObject coolingCoil = FindObject(
                    vrfTerminalUnit["Cooling Coil Object Type"].Value,
                    vrfTerminalUnit["Cooling Coil Object Name"].Value);

                coolingCoil["Coil Air Inlet Node"].Value = oldTerminal["Air Outlet Node Name"].Value;

                IdfObject zoneInletNodeList = FindObject("NodeList", zoneSpec.ZoneInletNodeListName);
                RemoveNodeFromNodeList(zoneInletNodeList, oldTerminal["Air Outlet Node Name"].Value);

                // Insert mixer then remove old terminal
                idf.Load(mixerIdfText);
                idf.Remove(oldTerminal);

                // Update VRF terminal unit inlet node reference to align with the new arrangement
                vrfTerminalUnit["Terminal Unit Air Inlet Node Name"].Value = oldTerminal["Air Outlet Node Name"].Value;
            }
        }

        private IdfObject FindObject(string objectType, string objectName)
        {
            try
            {
                return idf[objectType].First(c => c[0] == objectName);
            }
            catch
            {
                throw new MissingFieldException(
                    String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        // Remove a specific node name from a NodeList (if present)
        private void RemoveNodeFromNodeList(IdfObject nodeList, string nodeName)
        {
            for (int i = 0; i < nodeList.Fields.Count; i++)
            {
                if (nodeList[i].Value == nodeName)
                {
                    nodeList.RemoveField(i);
                    break;
                }
            }
        }

        // Holds the names of the relevant objects found for a given zone
        public struct ZoneEquipmentSpecification
        {
            public string ZoneEquipmentListName;
            public string DistributionUnitName;

            public string TerminalType;
            public string TerminalName;

            public string ZoneUnitType;
            public string ZoneUnitName;

            // Assumes NodeList naming convention derived from the equipment list name
            public string ZoneInletNodeListName
            {
                get
                {
                    if (ZoneEquipmentListName != null)
                        return ZoneEquipmentListName.Replace(" Equipment", " Air Inlet Node List");
                    return null;
                }
            }
        }

        // Scan ZoneHVAC:EquipmentList and return zones that contain both:
        // - a distribution unit whose air terminal type matches terminalType
        // - a zone unit object whose type matches zoneUnitType (VRF terminal unit)
        private List<ZoneEquipmentSpecification> FindZonesWithCavAndVrf(string terminalType, string zoneUnitType)
        {
            List<ZoneEquipmentSpecification> specifications = new List<ZoneEquipmentSpecification>();

            IEnumerable<IdfObject> allZoneEquipmentLists = idf["ZoneHVAC:EquipmentList"];
            foreach (IdfObject zoneEquipmentList in allZoneEquipmentLists)
            {
                string terminalName = null;
                string zoneUnitName = null;
                string distributionUnitName = null;

                bool includesTerminal = false;
                bool includesZoneUnit = false;

                for (int i = 0; i < zoneEquipmentList.Fields.Count - 1; i++)
                {
                    string fieldValue = zoneEquipmentList[i].Value;

                    // Detect distribution unit entry
                    if (String.Equals(fieldValue, DistributionUnitType, StringComparison.OrdinalIgnoreCase))
                    {
                        distributionUnitName = zoneEquipmentList[i + 1].Value;

                        IdfObject distributionUnit = FindObject(DistributionUnitType, distributionUnitName);
                        string airTerminalType = distributionUnit["Air Terminal Object Type"].Value;

                        if (String.Equals(airTerminalType, terminalType, StringComparison.OrdinalIgnoreCase))
                        {
                            terminalName = distributionUnit["Air Terminal Name"].Value;
                            includesTerminal = true;
                        }
                    }

                    // Detect VRF terminal unit entry
                    if (String.Equals(fieldValue, zoneUnitType, StringComparison.OrdinalIgnoreCase))
                    {
                        zoneUnitName = zoneEquipmentList[i + 1].Value;
                        includesZoneUnit = true;
                    }
                }

                if (includesTerminal && includesZoneUnit)
                {
                    specifications.Add(new ZoneEquipmentSpecification
                    {
                        ZoneEquipmentListName = zoneEquipmentList[0].Value,
                        DistributionUnitName = distributionUnitName,
                        TerminalType = terminalType,
                        TerminalName = terminalName,
                        ZoneUnitType = zoneUnitType,
                        ZoneUnitName = zoneUnitName
                    });
                }
            }

            return specifications;
        }
    }
}