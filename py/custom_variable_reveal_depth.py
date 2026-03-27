"""
Use the "Custom Script" variable to update the Reveal depth attribute.

The expected script variable name is the "Reveal depth".

"""
import ctypes


REVEAL_DEPTH_ATTR = "RevealOutsideProjection"


def show_message(title, text):
  ctypes.windll.user32.MessageBoxW(0, text, title, 0)


def on_design_variable_changed(opt_var_id, variable_current_value):
  site = api_environment.Site
  table = site.GetTable("OptimisationVariables")
  record = table.Records.GetRecordFromHandle(opt_var_id)
  opt_variable_name = record["Name"]

  if opt_variable_name == "Reveal depth":
      for zone in (zone for block in active_building.BuildingBlocks for zone in block.Zones):
          for surface in zone.Surfaces:
              for opening in (opening for opening in surface.Openings if str(opening.Type) == "Window"):
                   opening.SetAttribute(REVEAL_DEPTH_ATTR, str(variable_current_value))


