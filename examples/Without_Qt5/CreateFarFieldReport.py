# pyaedt v0.6.4
# This script plots Radiation patterns from several antennas and export them as csv files.
# The "project_name" variable must be adapted to your path.
import sys
import os
import pyaedt
from pyaedt import Hfss

sys.setrecursionlimit(100_000)
print(chr(27) + "[2J")

# inputs
project_name = "D:\\5-Customer_Issues\\Aptiv\\16-sept-2022\\simple_radar_module_2TX_3RX.aedt"
design_name = "0_radar_modul"

# Delete "lock" file
pyaedt.generic.general_methods.remove_project_lock(project_name)

# Open the hfss session
# h = pyaedt.Hfss(projectname= project_name, designname= design_name, solution_type = "Terminal", specified_version="2022.2", non_graphical=False, new_desktop_session=True) # h = pyaedt.Hfss(specified_version="2022.2", non_graphical=False)
h = pyaedt.Hfss(projectname= project_name, designname= design_name, specified_version="2022.2", non_graphical=False, new_desktop_session=True)
# h = pyaedt.Hfss(designname= design_name, specified_version="2022.2", non_graphical=False, new_desktop_session=False)

# Get the excitations name
ant_port_name = h.excitations
wd = h.project_path

# Activate the design
h.set_active_design(design_name)

# Delete the radiation 3D Sphere
try:
    h.oradfield.DeleteSetup("Sphere_Custom")
except:
    pass

# Create an infinite sphere for the radiation pattern
h.insert_infinite_sphere(definition='Theta-Phi', x_start=0, x_stop=180, x_step=2, y_start=-180, y_stop=180, y_step=2, units='deg',  custom_coordinate_system=None,  name="Sphere_Custom")

# Select parameter variation for the FarField plot
variations = h.available_variations.nominal_w_values_dict
variations["Freq"] = ["77GHz"]
variations["Theta"] = ["All"]
variations["Phi"] = ["All"]

# Set excitation amplitude to 1V
for i in range(len(ant_port_name)):
    h.edit_source(ant_port_name[i], "1V")

# Set the "Source context"
h.set_source_context(ant_port_name)

# Simulate the project
# h.analyze_setup("Setup1")

# Create the "3D polar plot Far Fields report"
for i, portname in enumerate(ant_port_name):
    context = {"Context": "Sphere_Custom", "SourceContext": ant_port_name[i]}
    h.post.create_report("db(GainTotal)", h.nominal_adaptive, variations=variations, primary_sweep_variable="Theta", plot_type="3D Polar Plot", context=context, report_category="Far Fields", plotname="Plot_Ant%d" %(i+1))
    h.post.export_report_to_csv(wd, "Plot_Ant%d" %(i+1))

# Save project and close HFSS
h.save_project()
h.release_desktop()
