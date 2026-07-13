# DB-scripts

This repository provides a comprehensive collection of scripts and plugins designed to extend the modeling and simulation capabilities of DesignBuilder and EnergyPlus. These tools leverage the DesignBuilder API and EnergyPlus Runtime Language (Erl) to automate workflows, implement custom control logic, and perform advanced post-processing.

## Overview of Scripting Types

There are three different languages to customize simulations:

1. **Energy Management System (EMS)**  
   EMS scripts provide high-level, supervisory control by overriding standard EnergyPlus behavior during the simulation. Using the EnergyPlus Runtime Language (Erl), these scripts connect Sensors (to get data) and Actuators (to set control results). Common applications include control of operation based on other variables, such as humidity-based window control.

2. **C# (CS-Script)**  
   C# scripts are used to modify the EnergyPlus Input Data File (IDF) immediately before simulation or to report on internal data at specific hook points. These scripts typically utilize the EpNet library to parse and manipulate IDF objects. Common applications include customizing outputs and adding HVAC and library components which are present in EnergyPlus but have not yet been integrated into DesignBuilder's interface.

3. **Python scripting**  
   Similar to C# scripts, Python scripts are primarily used to hook into specific stages of the simulation process. These scripts utilize the `api_environment` singleton object and the `eppy` scripting tool to access model attributes, site operations, and results tables.

## Integration and Workflow

**Access.** Scripts are managed via the Script Manager tool, where users can enable, edit, and compile code.

To add a script to your DesignBuilder model:

1. Click on the Scripts toolbar icon (available at building, block or zone levels).
2. Enable scripts in the General tab.
3. Click the Add new item button in the folder of your corresponding script type (EMS, CSharp, Python).
4. Paste the code and make sure that the Enable program option is active, then click OK to save the changes.
5. Confirm that the code is active by checking if it has a red circle in its icon in the Scripts screen.
6. Run your simulation.

**Storage.** Unlike plugins, scripts are saved directly within the DesignBuilder `.dsb` model file, ensuring portability with the project model.

**Execution.** Control is passed to scripts at specific hook points, such as `BeforeEnergySimulation` or `AfterEnergySimulation`, allowing for precise timing of model modifications. The full list of hooks available can be found in the DesignBuilder Extensibility User Guide.

> **Note:** If you have multiple active scripts and plugins implemented at the same hook, DesignBuilder will pass control in the following order: all scripts in the order they appear in the Script Manager, followed by all plugins in alphabetical order of the plugin’s assembly name.

## Technical References

1. [DesignBuilder Extensibility User Guide](https://designbuilder.co.uk/helpv2025.1/#Extensibility.htm)  
   Technical reference for the DesignBuilder API and plugin development.

2. [EMS Application Guide](https://energyplus.readthedocs.io/en/stable/ems-application-guide/ems-application-guide.html)  
   Detailed reference for Erl syntax and simulation calling points.

3. [EnergyPlus InputOutputReference](https://energyplus.net/assets/nrel_custom/pdfs/pdfs_v25.2.0/InputOutputReference.pdf)  
   Used for identifying exact fields and object structures for IDF manipulation.

## Disclaimer

These scripts are provided as-is without warranty. DesignBuilder takes no responsibility for simulation results, accuracy, or any issues arising from the use of this script. Users are responsible for validating all outputs and ensuring the script meets their specific modeling requirements.
