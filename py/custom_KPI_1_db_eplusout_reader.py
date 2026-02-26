"""
Create a basic custom KPI script to read the results using db_eplusout_reader library.

Detailed information:
https://designbuilder.co.uk/helpv7.2/#PythonScriptingExample2.htm

"""

from db_eplusout_reader import Variable, get_results
from db_eplusout_reader.constants import RP


# ensure that SQLite outputs are generated
def before_energy_idf_generation():
    site = api_environment.Site
    for building in site.Buildings:
        building.SetAttribute("SSSQLiteOP", "1")


def after_energy_simulation():
    file = api_environment.EnergyPlusFolder + r"eplusout.sql"
    variables = [
        Variable("", "Heating Coil Electricity Energy", "J"),
        Variable("", "Cooling Coil Electricity Energy", "J"),
        Variable("", "Fan Electricity Rate", "W"),
    ]
    results = get_results(file, variables=variables, alike=True, frequency=RP)
    convertJ_kWh = 3600000
    convertW_kWh = 8.76

    heating = results.arrays[0][0] / convertJ_kWh
    cooling = results.arrays[1][0] / convertJ_kWh
    fan = results.arrays[2][0] * convertW_kWh

    system_energy_kWh = str(heating + cooling + fan)

    site = api_environment.Site
    table = site.GetTable("ParamResultsTmp")

    record = table.AddRecord()
    record[0] = "Total system energy (kWh)"
    record[1] = str(system_energy_kWh)
