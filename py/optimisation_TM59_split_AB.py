"""
Use CIBSE TM59 results as optimisation KPI.

Criteria A & B are reported as two independent output variables.

Required model inputs:
- Enable TM59 outputs
- Enable SQLite output
- Add two "Custom Script" KPIs:
        "TM59 Discomfort Crit A %"
        "TM59 Discomfort Crit B %"

- Run a "Standard" simulation to make sure that TM59 outputs are present
        (temporarily enable # show message... debugging messages)

The script calculates a percentage of failing TM59 relevant zones.

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


def zone_fails_a(zone_name, outputs_path):
    """Verify if given zone meets tm59 criteria."""
    variable1 = Variable("EMS", "CIBSE TM59 Criterion A {}".format(zone_name), "%")

    criterion_a_results = get_results(outputs_path, variable1, frequency=RP)
    criterion_a = criterion_a_results.scalar
    fail = criterion_a > THRESHOLD_A

    return fail


def zone_fails_b(zone_name, outputs_path):
    """Verify if given zone meets tm59 criteria."""
    variable2 = Variable("EMS", "CIBSE TM59 Criterion B {}".format(zone_name), "H")

    criterion_b_results = get_results(outputs_path, variable2, frequency=RP)
    criterion_b = criterion_b_results.scalar
    fail = criterion_b > THRESHOLD_B

    return fail


def after_energy_simulation():
    site = api_environment.Site
    fail_area_a, fail_area_b = 0, 0
    total_area_a, total_area_b = 0, 0  # area only includes TM59 zones
    outputs_path = os.path.join(api_environment.EnergyPlusFolder, "eplusout.sql")

    for block in active_building.BuildingBlocks:
        for zone in block.Zones:
            zone_name = zone.GetAttribute("SSEPObjectNameInOP")
            try:
                if zone_fails_a(zone_name, outputs_path):
                    fail_area_a += zone.FloorArea
                total_area_a += zone.FloorArea
            except NoResults:
                continue
            try:
                if zone_fails_b(zone_name, outputs_path):
                    fail_area_b += zone.FloorArea
                total_area_b += zone.FloorArea
            except NoResults:
                pass
    try:
        percentage_area_fail_a = (fail_area_a / total_area_a) * 100
    except ZeroDivisionError:
        raise NoResults("No TM59 results found!")

    try:
        percentage_area_fail_b = (fail_area_b / total_area_b) * 100
    except ZeroDivisionError:
        show_message("Results TM59", "No criterion B results")
        percentage_area_fail_b = 0

    # show_message("Results TM59", "{}% fails criterion A, total_area {}m2, fail_area: {}m2".format(percentage_area_fail_a, total_area_a, fail_area_a))
    # show_message("Results TM59", "{}% fails criterion B, total_area {}m2, fail_area: {}m2".format(percentage_area_fail_b, total_area_b, fail_area_b))

    table = site.GetTable("ParamResultsTmp")
    record = table.AddRecord()
    record[0] = "TM59 Discomfort Crit A %"
    record[1] = str(percentage_area_fail_a)

    record = table.AddRecord()
    record[0] = "TM59 Discomfort Crit B %"
    record[1] = str(percentage_area_fail_b)
