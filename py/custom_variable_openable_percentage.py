"""
Use the "Custom Script" variable to update the % Glazing Area Opens attribute.

Openings with the original value equal to "0" are ignored.

The expected script variable name is the "Glazing opening percentage".

"""
import ctypes


PERC_GLAZING_ATTR = "ExtWinNaturalVentilationPercOpeningValue"


def show_message(title, text):
   ctypes.windll.user32.MessageBoxW(0, text, title, 0)


def on_design_variable_changed(opt_var_id, variable_current_value):
   site = api_environment.Site
   table = site.GetTable("OptimisationVariables")
   record = table.Records.GetRecordFromHandle(opt_var_id)
   opt_variable_name = record["Name"]

   show_message("FOO", type(variable_current_value))

   if opt_variable_name == "Glazing opening percentage":
       for zone in (zone for block in active_building.BuildingBlocks for zone in block.Zones):
           for surface in zone.Surfaces:
               for opening in (opening for opening in surface.Openings if str(opening.Type) == "Window"):
                   original_value = opening.GetAttributeAsDouble(PERC_GLAZING_ATTR)
                   if original_value > 0:
                        opening.SetAttribute(PERC_GLAZING_ATTR, str(variable_current_value))