/*
Check Proposed vs Baseline (ASHRAE 90.1) for the Active Building

Purpose
- Determine whether the currently simulated building is the ASHRAE 90.1 Proposed case or Baseline case.
- NOTE: This can be used to conditionally apply other script actions only to specific building variants.

How to Use

Configuration
- Attribute key checked: "ASHRAE901Type"
- Proposed value expected: "1-Proposed"

Prerequisites
- The model must be configured to to an ASHRAE 90.1 Model (Proposed/ Baseline vesrions)

DISCLAIMER: This script is provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
*/

using System.Windows.Forms;
using System;
using DB.Api;
using DB.Extensibility.Contracts;

namespace DB.Extensibility.Scripts
{
    public class CheckProposed : ScriptBase, IScript
    {
        private bool IsCurrentProposed()
        {
            // Read ASHRAE 90.1 case type from the Active Building attribute
            string ashrae901Type = ActiveBuilding.GetAttribute("ASHRAE901Type");
            return ashrae901Type == "1-Proposed";
        }

        public override void BeforeEnergyIdfGeneration()
        {
            // Hook executed before EnergyPlus IDF generation; use this to gate other script actions
            if (IsCurrentProposed())
            {
                MessageBox.Show("Proposed");
            }
            else
            {
                MessageBox.Show("Baseline");
            }
        }
    }
}