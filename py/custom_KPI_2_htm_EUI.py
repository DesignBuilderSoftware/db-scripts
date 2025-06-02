"""
Create a basic custom KPI script using eppy html reader to read the results.

Detailed information:
https://designbuilder.co.uk/helpv8.0/#CustomKPIsExample.htm

"""

import ctypes
from os import path
from eppy.results import fasthtml


def show_message(title, text):
    ctypes.windll.user32.MessageBoxW(0, text, title, 0)


def after_energy_simulation():
    table_path = api_environment.EnergyPlusFolder + "eplustbl.htm"
    if path.exists(table_path):
        with open(table_path, "r") as filehandle:
            eui_table = fasthtml.tablebyname(filehandle, "EAp2-17a. Energy Use Intensity - Electricity")
            eui_table_content = eui_table[1]
            eui = str(eui_table_content[8][1])

            site = api_environment.Site
            table = site.GetTable("ParamResultsTmp")

            record = table.AddRecord()
            record[0] = "EUI (kWh/m2)"
            record[1] = eui
    else:
        show_message("Information", "File does not exist " + table_path)
