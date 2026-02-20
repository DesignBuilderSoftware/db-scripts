"""
Applying WWR variable in the "Calculated" natural ventilation mode
changes the openable area (as the area is represented as a percentage).

There may be cases when this approach is not suitable.

This script  modifies the "% Glazing Area Opens" for each WWR - to keep the
openable area constant (default 1m2).

"""

import ctypes


def show_message(title, text):
    ctypes.windll.user32.MessageBoxW(0, text, title, 0)


def before_energy_idf_generation():
    # define hard-set openable area [m2]
    openable_area = 1.0
    percantage_area_opens = "ExtWinNaturalVentilationPercOpeningValue"
    wwr = active_building.GetAttribute("PercGlazing")
    openings_summary = "Openings Summary:\n"
    openings_summary += "-" * 25

    show_message("WWR", "Building WWR: " + str(wwr) + "%")
    for block in active_building.BuildingBlocks:
        for zone in block.Zones:
            for surface in zone.Surfaces:
                for adjacency in surface.Adjacencies:
                    for opening in adjacency.Openings:
                        area = opening.Area
                        openable_percentage = str(min(openable_area / area * 100, 100))
                        opening.SetAttribute(percantage_area_opens, openable_percentage)
                        name = opening.GetAttribute("SSEPObjectNameInOP")
                        openings_summary += (
                            "\nOpening: "
                            + name
                            + "\n\t"
                            + "Area: "
                            + str(area)
                            + "\n\t"
                            + "Percentage: "
                            + openable_percentage
                        )
    show_message("Summary", openings_summary)
