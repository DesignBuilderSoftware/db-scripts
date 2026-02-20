/* 
Check if currently simulated buidling is Proposed or Baseline.

This can be used apply a script only to specific building.

*/

using System.Collections.Generic;
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
            string ashraeType = ActiveBuilding.GetAttribute("ASHRAE901Type");
            return ashraeType == "1-Proposed";
        }

        public override void BeforeEnergyIdfGeneration()
        {
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