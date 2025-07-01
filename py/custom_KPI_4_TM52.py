"""
Use CIBSE TM52 results as optimisation KPI.

Required model inputs:
    - Enable TM52 outputs
    - Enable SQLite output
    - Add a "Custom Script" KPI, use "TM52 Discomfort %" name
    - Run a "Standard" simulation to make sure that TM52 outputs are present

The script calculates a percentage of failing TM52 relevant zones.

Detailed information:
https://designbuilder.co.uk/helpv7.2/#CustomEMSScriptExample4.htm
"""

import ctypes
import os

from db_eplusout_reader import Variable, get_results
from db_eplusout_reader.constants import *
from db_eplusout_reader.exceptions import NoResults

threshold_crit1 = 3
threshold_crit2 = 6
threshold_crit3 = 0


def show_message(title, text):
    ctypes.windll.user32.MessageBoxW(0, text, title, 0)


def zone_fails_tm52(zone_name):
    """Verify if given zone meets tm52 criteria."""
    outputs_path = os.path.join(api_environment.EnergyPlusFolder, "eplusout.sql")
    variable1 = Variable("EMS", "CIBSE TM52 Criterion 1 {}".format(zone_name), "%")
    variable2 = Variable("EMS", "CIBSE TM52 Criterion 2 {}".format(zone_name), "C")
    variable3 = Variable("EMS", "CIBSE TM52 Criterion 3 {}".format(zone_name), "hr")

    criterion_1_results = get_results(outputs_path, variable1, frequency=RP)
    criterion_2_results = get_results(outputs_path, variable2, frequency=D)
    criterion_3_results = get_results(outputs_path, variable3, frequency=RP)

    criterion_1 = criterion_1_results.scalar
    crit2_series = criterion_2_results.first_array
    criterion_2 = max(crit2_series)
    criterion_3 = criterion_3_results.scalar

    counter = 0

    if criterion_1 > threshold_crit1:
        counter += 1

    if criterion_2 > threshold_crit2:
        counter += 1

    if criterion_3 > threshold_crit3:
        counter += 1

    return counter >= 2


def after_energy_simulation():
    site = api_environment.Site
    fail_area = 0
    total_area = 0  # area only includes TM52 zones

    for block in active_building.BuildingBlocks:
        for zone in block.Zones:
            zone_name = zone.GetAttribute("SSEPObjectNameInOP")
            try:
                if zone_fails_tm52(zone_name):
                    fail_area += zone.FloorArea
                total_area += zone.FloorArea
            except NoResults:
                pass

    try:
        percentage_area_fail = (fail_area / total_area) * 100
    except ZeroDivisionError:
        raise NoResults("No TM52 results found!")

    show_message("Results TM52", "{}% fails, total_area {}m2, fail_area: {}m2".format(percentage_area_fail, total_area, fail_area))

    table = site.GetTable("ParamResultsTmp")
    record = table.AddRecord()
    record[0] = "TM52 Discomfort %"
    record[1] = str(percentage_area_fail)
