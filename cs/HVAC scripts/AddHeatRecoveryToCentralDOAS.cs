/*
Central Dedicated Outdoor Air System (DOAS) Heat Recovery Injection Script

Purpose:
This DesignBuilder C# script finds an existing DOAS created via the DesignBuilder interface
and inserts a HeatExchanger:AirToAir:SensibleAndLatent into that DOAS outdoor air system.

Main Steps:
1) Find the existing AirLoopHVAC:DedicatedOutdoorAirSystem by name
2) Find its linked AirLoopHVAC:OutdoorAirSystem and EquipmentList
3) Identify the current first component in the OA equipment list
4) Insert a heat recovery heat exchanger before that first component
5) Rewire that first component inlet node to the HX supply outlet node
NOTE: The Heat Recovery device inserted as the first component after OA intake, before the coils (if exisitng) and fan.

How to Use:

Configuration
- The user should add the name of the DOAS loop which requires the Heat Recovery to be placed.
- Settings for the heat recovery are specified in the boilerplate HeatExchanger:AirToAir:SensibleAndLatent (adjust if required)
NOTE: The availability schedule for heat recovery follows the one of the DOAS equipment

Prerequisites / Placeholders
- A DOAS Loop has to be present in the DesignBuilder model (with required coils)
- Additional setings can be adjusted in the boilerplate HeatExchanger:AirToAir:SensibleAndLatent

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility
for simulation results, accuracy, or any issues arising from the use of this script.
Users are responsible for validating all outputs and ensuring the script meets their specific
modeling requirements.
*/

using System;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        public override void BeforeEnergySimulation()
        {
            IdfReader idf = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath
            );

            DoasIdfHandler doasIdfHandler = new DoasIdfHandler(idf);

            // ---------------------------
            // USER CONFIGURATION SECTION
            // ---------------------------
            // Existing loop name for the DOAS to be modified as it appears in the IDF.
            string doasLoopName = "DOAS Loop";
            DoasSpecs doas1 = new DoasSpecs(doasLoopName);
            doasIdfHandler.LoadDoas(doas1);
        }
    }

    public class DoasSpecs
    {
        public string DoasLoopName;

        public DoasSpecs() { }

        public DoasSpecs(string doasLoopName)
        {
            DoasLoopName = doasLoopName;
        }

        public string HxName { get { return this.DoasLoopName + " Heat Recovery Device"; } }
        public string HxSupplyOutletNode { get { return this.DoasLoopName + " Heat Recovery Supply Outlet Node"; } }
        public string HxReliefOutletNode { get { return this.DoasLoopName + " Heat Recovery Relief Outlet Node"; } }

        public string GetInfo()
        {
            string text = @"DOAS loop: {0}";
            return String.Format(text, this.DoasLoopName);
        }

        public string GetIDFObjects(string availabilityScheduleName, string outdoorAirInletNode, string mixerOutletNode)
        {
            // ---------------------------
            // USER CONFIGURATION SECTION
            // ---------------------------
            // Adjust parameters for the Heat Exchanger if required
            string idfObjects = @"
HeatExchanger:AirToAir:SensibleAndLatent,
   {0},                                                           !- Name
   {1},                                                           !- Availability Schedule Name
   autosize,                                                      !- Nominal Supply Air Flow Rate (m3/s)
   0.75,                                                          !- Sensible Effectiveness at 100% Heating Air Flow
   0.00,                                                          !- Latent Effectiveness at 100% Heating Air Flow
   0.75,                                                          !- Sensible Effectiveness at 75% Heating Air Flow
   0.00,                                                          !- Latent Effectiveness at 75% Heating Air Flow
   0.75,                                                          !- Sensible Effectiveness at 100% Cooling Air Flow
   0.00,                                                          !- Latent Effectiveness at 100% Cooling Air Flow
   0.75,                                                          !- Sensible Effectiveness at 75% Cooling Air Flow
   0.00,                                                          !- Latent Effectiveness at 75% Cooling Air Flow
   {2},                                                           !- Supply Air Inlet Node Name
   {3},                                                           !- Supply Air Outlet Node Name
   {4},                                                           !- Exhaust Air Inlet Node Name
   {5},                                                           !- Exhaust Air Outlet Node Name
   0.000,                                                         !- Nominal Electric Power (W)
   No,                                                            !- Supply Air Outlet Temperature Control
   Plate,                                                         !- Heat Exchanger Type
   None,                                                          !- Frost Control Type
   1.70,                                                          !- Threshold Temperature (C)
   0.167,                                                         !- Initial Defrost Time Fraction
   0.0240,                                                        !- Rate of Defrost Time Fraction Increase
   Yes;                                                           !- Economizer Lockout";

            return String.Format(
                idfObjects,
                this.HxName,
                availabilityScheduleName,
                outdoorAirInletNode,
                this.HxSupplyOutletNode,
                mixerOutletNode,
                this.HxReliefOutletNode
            );
        }
    }

    public class DoasIdfHandler
    {
        public IdfReader Reader;
        public DoasIdfHandler() { }
        public DoasIdfHandler(IdfReader idfReader)
        {
            Reader = idfReader;
        }

        public IdfObject FindObject(string objectType, string objectName)
        {
            try
            {
                return this.Reader[objectType].First(c => c[0] == objectName);
            }
            catch (Exception)
            {
                throw new Exception(String.Format("Cannot find object: {0}, type: {1}", objectName, objectType));
            }
        }

        // Returns the field index of the air inlet node for the first OA component.
        private int GetAirInletFieldIndex(string objectType)
        {
            if (objectType == "Coil:Cooling:Water")
            {
                return 11;
            }

            if (objectType == "Coil:Heating:Water")
            {
                return 6;
            }

            if (objectType == "Fan:SystemModel")
            {
                return 2;
            }

            throw new Exception(String.Format("Unsupported first OA component type: {0}", objectType));
        }

        public void LoadDoas(DoasSpecs doasSpecs)
        {
            IdfObject doas = FindObject("AirLoopHVAC:DedicatedOutdoorAirSystem", doasSpecs.DoasLoopName);

            // The DOAS points to the OA system and mixer that need to be edited.
            string outdoorAirSystemName = doas[1].Value;
            string availabilityScheduleName = doas[2].Value;
            string mixerName = doas[3].Value;

            IdfObject outdoorAirSystem = FindObject("AirLoopHVAC:OutdoorAirSystem", outdoorAirSystemName);
            IdfObject mixer = FindObject("AirLoopHVAC:Mixer", mixerName);

            string equipmentListName = outdoorAirSystem[2].Value;
            IdfObject equipmentList = FindObject("AirLoopHVAC:OutdoorAirSystem:EquipmentList", equipmentListName);

            // The first component in the list is the current first downstream OA component (HX will be inserted before it).
            string firstComponentType = equipmentList[1].Value;
            string firstComponentName = equipmentList[2].Value;

            IdfObject firstComponent = FindObject(firstComponentType, firstComponentName);

            int inletFieldIndex = GetAirInletFieldIndex(firstComponentType);

            string originalFirstComponentInletNode = firstComponent[inletFieldIndex].Value;
            string mixerOutletNode = mixer[1].Value;

            string hxIdfObjects = doasSpecs.GetIDFObjects(
                availabilityScheduleName,
                originalFirstComponentInletNode,
                mixerOutletNode
            );

            // Add the new HX object to the IDF.
            this.Reader.Load(hxIdfObjects);

            // Insert the HX as the first item in the OA equipment list.
            equipmentList.InsertField(1, "HeatExchanger:AirToAir:SensibleAndLatent");
            equipmentList.InsertField(2, doasSpecs.HxName);

            // Reconnect the original first component so it now receives air from the HX outlet.
            firstComponent[inletFieldIndex].Value = doasSpecs.HxSupplyOutletNode;

            this.Reader.Save();
        }
    }
}