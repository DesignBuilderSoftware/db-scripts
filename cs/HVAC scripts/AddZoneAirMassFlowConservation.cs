/*
Adds ZoneAirMassFlowConservation to the simulation IDF so that interzone mixing flows are balanced through the zone return air nodes.

Purpose:
DesignBuilder writes the interzone airflow path as a ZoneMixing object, which is a one-way flow: it enters the energy and mass balance of the receiving zone
only, and the source zone is unaffected. 
This script adds the ZoneAirMassFlowConservation object, so that EnergyPlus recalculates each zone's return air flow to close the balance.

With "AdjustReturnOnly", the zone total return air mass flow rate becomes:
    mR = MAX(0, mS - [mEX,tot - mEXF,bal] + [mXR - mXS])
where mS is the zone supply, mXR the mixing received and mXS the mixing sent.
The source zone therefore returns less by the transfer rate, and the receiving
zone returns more by the same amount, while the air loop total is unchanged.

Main Steps:
1) Insert ZoneAirMassFlowConservation for return-side adjustment only
2) Save the modified IDF so EnergyPlus uses it

How to Use:

Configuration

- AdjustMode controls field A1. 
  "AdjustReturnOnly" adjusts the return flows and leaves the mixing flows as entered.
  The other documented keys are AdjustMixingOnly, AdjustMixingThenReturn, AdjustReturnThenMixing and None.
- InfiltrationMethod is set to "None" so that infiltration is excluded from the mass balance calculation. 
  Use AddInfiltrationFlow or AdjustInfiltrationFlow only if infiltration is intended to make up the difference.
- Infiltration Balancing Zones is deliberately left blank as it is not used when InfiltrationMethod is None.

Prerequisites / Placeholders
- The interzone airflow path should be entered through the DB interface (Tools > Interzone Airflow).
- At least one ZoneMixing object must exist, otherwise the object has no effect. 
  ZoneAirMassFlowConservation applies only to controlled zones (zones with a ZoneHVAC:EquipmentConnections object) that also have a zone
  mixing or infiltration object

Notes:
- This object is unique and model-wide, not per zone. 
  It applies to every zone in the model, although zones in an air loop with no mixing objects and no exhaust fans are always balanced and so are unaffected.
- Only the BeforeEnergySimulation hook is used, so the DesignBuilder Heating and Cooling design calculations do not see this object. 
  Uncomment the two hooks at the bottom if the transfer air should also be present during sizing.

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues
arising from the use of this script. Users are responsible for validating all outputs and ensuring the script meets their specific modelling requirements.
*/

using System.Runtime;
using System;
using System.Linq;
using DB.Extensibility.Contracts;
using DB.Api;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class AddZoneAirMassFlowConservation : ScriptBase, IScript
    {
        // ----------------------------
        // USER CONFIGURATION SECTION
        // ----------------------------

        // Field A1: Adjust Zone Mixing and Return For Air Mass Flow Balance
        private const string AdjustMode = "AdjustReturnOnly";
        // Field A2: Infiltration Balancing Method
        private const string InfiltrationMethod = "None";

        public override void BeforeEnergySimulation()
        {
            ModifyIdf();
        }

        // Uncomment to apply the object to the design calculations as well.
        // public override void BeforeHeatingSimulation() { ModifyIdf(); }
        // public override void BeforeCoolingSimulation() { ModifyIdf(); }

        private void ModifyIdf()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            // ZoneAirMassFlowConservation is a unique object. Adding a second one is a fatal error, so check before inserting.
            bool alreadyPresent =
                idfReader["ZoneAirMassFlowConservation"].ToList().Count > 0;

            if (!alreadyPresent)
            {
                idfReader.Load(
                    "ZoneAirMassFlowConservation, " +
                    AdjustMode + ", " +          // Adjust mixing and return
                    InfiltrationMethod + ", " +  // Infiltration balancing method
                    ";");                        // Infiltration balancing zones
            }

            idfReader.Save();
        }
    }
}
