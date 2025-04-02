"""
Use the "Custom Script" variable to update the maximum mechanical ventilation ach value.

The script retrieves variable values from two custom script variables:
    - Bedroom ach
    - Living ach

Target zones can be filtered by a set of substrings, for instance, the current variable value
for the "Living ach" is applied to all zones with "living" or "lkd" in the zone title.

"""


MECH_VENT_ACH_ATTR = "MechanicalVentilationValue"


def filter_zones(building, *args):
    """ Retrieve zones within a building where the zone title contains a specific key. """
    for zone in (zone for block in building.BuildingBlocks for zone in block.Zones):
        zone_name = zone.GetAttribute("title")
        if any(key.lower() in zone_name.lower() for key in args):
            yield zone


def on_design_variable_changed(opt_var_id, variable_current_value):
    site = api_environment.Site
    table = site.GetTable("OptimisationVariables")
    record = table.Records.GetRecordFromHandle(opt_var_id)
    opt_variable_name = record["Name"]

    if opt_variable_name == "Bedroom ach":
        for zone in filter_zones(active_building, "bedroom"):
            zone.SetAttribute(MECH_VENT_ACH_ATTR, str(variable_current_value))

    elif opt_variable_name == "Living ach":
        for zone in filter_zones(active_building, "living", "lkd"):
            zone.SetAttribute(MECH_VENT_ACH_ATTR, str(variable_current_value))
