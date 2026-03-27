"""
Create a custom KPI script using db_eplusout_reader library to read the results.

This script should be applied in combination with Discomfort Area EMS script.

Detailed information:
https://designbuilder.co.uk/helpv7.2/#PythonScriptingExample3.htm

"""

from db_eplusout_reader import Variable, get_results
from db_eplusout_reader.constants import *


# ensure that SQLite outputs are generated
def before_energy_idf_generation():
    site = api_environment.Site
    for building in site.Buildings:
        building.SetAttribute("SSSQLiteOP", "1")


def after_energy_simulation():
    file = api_environment.EnergyPlusFolder + r"eplusout.sql"
    variables = Variable("EMS", "Percentage Discomfort Area", "")
    results = get_results(file, variables=variables, alike=True, frequency=RP)

    discomfort_hrs = results.scalar
    site = api_environment.Site
    table = site.GetTable("ParamResultsTmp")

    record = table.AddRecord()
    record[0] = "PPD Discomfort > 20%"
    record[1] = str(discomfort_hrs)
