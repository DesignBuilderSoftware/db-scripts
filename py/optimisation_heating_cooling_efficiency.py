"""
This script demonstrates how to apply script design variables by updating the IDF file.

It requires 'Detailed HVAC' mode with specific components in place:
- Boiler:HotWater
- Chiller:Electric:EIR

One of the suitable templates is 'FCU 4-pipe, Air-cooled chiller'.


Actuated variables are 'boiler nominal efficiency' and 'chiller reference COP'.

The range is defined via script 'Design Variable' with specific names:
- 'HeatingEff' (boiler nominal efficiency)
- 'CoolingCOP' (chiller reference COP)

"""

import ctypes

from eppy import modeleditor
from eppy.modeleditor import IDF


def show_message(title, text):
    ctypes.windll.user32.MessageBoxW(0, text, title, 0)


def before_energy_simulation():
    IDF.setiddname(api_environment.EnergyPlusInputIddPath)
    idf_file = IDF(api_environment.EnergyPlusInputIdfPath)

    # extract values from optimisation table
    site = api_environment.Site
    table = site.GetTable("OptimisationVariables")

    boilers = idf_file.idfobjects["Boiler:HotWater".upper()]
    heating_eff_row = table.Records["HeatingEff"]

    heating_eff = heating_eff_row["VariableCurrentValue"]
    if not heating_eff:
        # use this in standard simulation
        heating_eff = "0.95"

    for boiler in boilers:
        if heating_eff != "UNKNOWN":
            boiler.Nominal_Thermal_Efficiency = heating_eff
        else:
            show_message(
                "ERROR",
                "Cannot set heating COP, unknown value in table OptimisationVariables",
            )

    chillers = idf_file.idfobjects["Chiller:Electric:EIR".upper()]
    cooling_cop_row = table.Records["CoolingCOP"]
    cooling_cop = cooling_cop_row["VariableCurrentValue"]
    if not cooling_cop:
        # use this in standard simulation
        cooling_cop = "3.5"

    for chiller in chillers:
        if cooling_cop != "UNKNOWN":
            chiller.Reference_COP = cooling_cop
        else:
            show_message(
                "ERROR",
                "Cannot set cooling COP, unknown value in table OptimisationVariables",
            )

    idf_file.save()
