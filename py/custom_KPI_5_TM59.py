"""
Use CIBSE TM59 results as optimisation KPI.

Required model inputs:
    - Enable TM59 outputs
    - Enable SQLite output
    - Add a "Custom Script" KPI, use "TM59 Discomfort %" name
    - Run a "Standard" simulation to make sure that TM59 outputs are present

 The script calculates a percentage of failing TM59 relevant zones.

Detailed information:
https://designbuilder.co.uk/helpv7.2/#CustomEMSScriptExample5.htm

"""

import ctypes
import os

from db_eplusout_reader import Variable, get_results
from db_eplusout_reader.constants import *
from db_eplusout_reader.exceptions import NoResults

THRESHOLD_A = 3
THRESHOLD_B = 32


def show_message(title, text):
    ctypes.windll.user32.MessageBoxW(0, text, title, 0)


def zone_fails_tm59(zone_name):
    """Verify if given zone meets tm59 criteria."""
    outputs_path = os.path.join(api_environment.EnergyPlusFolder, "eplusout.sql")
    variable1 = Variable("EMS", "CIBSE TM59 Criterion A {}".format(zone_name), "%")
    variable2 = Variable("EMS", "CIBSE TM59 Criterion B {}".format(zone_name), "H")

    criterion_a_results = get_results(outputs_path, variable1, frequency=RP)
    criterion_a = criterion_a_results.scalar
    fail = criterion_a > THRESHOLD_A

    try:
        criterion_b_results = get_results(outputs_path, variable2, frequency=RP)
        criterion_b = criterion_b_results.scalar
        fail = fail or criterion_b > THRESHOLD_B
    except NoResults:
        pass

    return fail


def after_energy_simulation():
    site = api_environment.Site
    fail_area = 0
    total_area = 0  # area only includes TM59 zones

    for block in active_building.BuildingBlocks:
        for zone in block.Zones:
            zone_name = zone.GetAttribute("SSEPObjectNameInOP")
            try:
                if zone_fails_tm59(zone_name):
                    fail_area += zone.FloorArea
                total_area += zone.FloorArea
            except NoResults:
                pass

    try:
        percentage_area_fail = (fail_area / total_area) * 100
    except ZeroDivisionError:
        raise NoResults("No TM59 results found!")

    show_message("Results TM59", "{}% fails, total_area {}m2, fail_area: {}m2".format(percentage_area_fail, total_area, fail_area))

    table = site.GetTable("ParamResultsTmp")
    record = table.AddRecord()
    record[0] = "TM59 Discomfort %"
    record[1] = str(percentage_area_fail)
