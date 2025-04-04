"""
Use the "Custom Script" variable to update the g-value for specified templates.

The script retrieves variable values from two custom script variables:
   - South facade g
   - East facade g

The typical case is the orientation-specific g-value, i.e. the "East facade glazing"
glazing type could be applied to all eastern windows, "South facade glazing" to all
windows facing south.

"""


GLAZING_TEMPLATE_G_VALUE_ATTR = "TransSolar"


def find_record_by_name(table, name):
   """Retrieve the table record by caseless name look-up."""
   lowercase_name = name.lower()
   for row in table.Records:
       if row["Name"].lower() == lowercase_name:
           return row
   raise KeyError("Cannot find {0} in table {1}".format(name, table))


def on_design_variable_changed(opt_var_id, variable_current_value):
   site = api_environment.Site
   glazing_table = site.GetTable("Glazing")

   opt_table = site.GetTable("OptimisationVariables")
   row = opt_table.Records.GetRecordFromHandle(opt_var_id)
   opt_variable_name = row["Name"]

   if opt_variable_name == "South facade g":
       row = find_record_by_name(glazing_table, "South facade glazing")
       row[GLAZING_TEMPLATE_G_VALUE_ATTR] = variable_current_value

   elif opt_variable_name == "East facade g":
       row = find_record_by_name(glazing_table, "East facade glazing")
       row[GLAZING_TEMPLATE_G_VALUE_ATTR] = variable_current_value

