# Tasks:
# Simulate the "1_PCB_Ex_4Ports" project, then generate a spice file of the project (Right Click "Setup1" --> Click "Network Data Explorer", select all S-Parameters (Setup1:Sweep)
# and then click "Export broadband model to file" with following options :
# HSPICE, Desired filtering Error = 0.01%, Maximum Order = 10000, TWA

# In Circuit :
# 1-Place the "SPICE component"
# 2-Simulate thw circuit
# Then execute the script

import sys
import os
import pyaedt
from pyaedt import Circuit
from pyaedt import Hfss

sys.setrecursionlimit(100_000)
print(chr(27) + "[2J")

# inputs
project_name = "X:\\...\\test_hfss_circuit_spice.aedt" # Modify accordingly

no_plots = False
cir_design_name = "2_Circuit_SPICE_LINK_4P_Spice"
hfss_design_name = "2_PCB_Ex_4Ports_spice"
spice_id = 65517

# Delete "lock" file
pyaedt.generic.general_methods.remove_project_lock(project_name)

# Open the "Circuit" and "HFSS" sessions
cir = pyaedt.Circuit(designname=cir_design_name, non_graphical=False, specified_version="2022.2")
h = pyaedt.Hfss(designname=hfss_design_name, non_graphical=False, specified_version="2022.2")

# Get the "working directory"
wd = cir.project_path

# Find the "SPICE component" --> create a "V-probe" at each port
npins = cir.modeler.components[spice_id].pins.__len__() # cir.modeler.components.get_pins(id).__len__()
probe = [None] * npins

# Delete the VPROBE components
for id in cir.modeler.components.components:
    name = cir.modeler.components.components[id].composed_name
    try:
        if 'VPROBE' in name:
            cir.layouteditor.Delete(id)
    except:
        pass
    try:
        if 'PagePort' in name:
            cir.oeditor.Delete(["NAME:Selections", "Selections:=", [name]])
    except :
        pass

# Add new VPROBE components
for i in range(npins):
    probe[i] = cir.modeler.components.components_catalog["Probes:VPROBE"].place("bla" + str(i))
    probe[i].parameters["Name"] = cir.modeler.components.components[spice_id].pins[i].name
    probe[i].pins[0].connect_to_component(cir.modeler.components.components[spice_id].pins[i])

# Simulate the project
cir.analyze_setup("NexximTransient")

# Delete all ".tab" files from wd
for file in os.listdir(wd):
    if file.endswith(".tab"):
        os.remove(os.path.join(wd, file))

# Delete all reports
if len(cir.post.all_report_names) != 0:
    cir.oreportsetup.DeleteAllReports()

# Create reports and ".tab" files in wd
for i in range(int(npins/2)):
    # Create plots and save data to a file
    new_report = cir.post.reports_by_category.spectral("mag(V(Port%d_T1))" %(i+1))
    new_report.window = "Hanning"
    new_report.max_frequency = "100MHz"
    new_report.time_start = "1us"
    new_report.time_stop = "11us"
    new_report.create()
    new_report.edit_x_axis_scaling(units='Hz')
    cir.post.rename_report(new_report.plot_name, "Port%d_T1_mag" %(i+1))
    cir.post.export_report_to_file(wd, "Port%d_T1_mag" %(i+1), ".tab")

    new_report = cir.post.reports_by_category.spectral("ang_rad(V(Port%d_T1))" %(i+1))
    new_report.window = "Hanning"
    new_report.max_frequency = "100MHz"
    new_report.time_start = "1us"
    new_report.time_stop = "11us"
    new_report.create()
    new_report.edit_x_axis_scaling(units='Hz')
    new_report.edit_x_axis_scaling(linear_scaling=False)
    cir.post.rename_report(new_report.plot_name, "Port%d_T1_ang" %(i+1))
    cir.post.export_report_to_file(wd, "Port%d_T1_ang" %(i+1), ".tab")

    new_report = cir.post.reports_by_category.spectral("mag(V(Port%d_T2))" %(i+1))
    new_report.window = "Hanning"
    new_report.max_frequency = "100MHz"
    new_report.time_start = "1us"
    new_report.time_stop = "11us"
    new_report.create()
    new_report.edit_x_axis_scaling(units='Hz')
    cir.post.rename_report(new_report.plot_name, "Port%d_T2_mag" %(i+1))
    cir.post.export_report_to_file(wd, "Port%d_T2_mag" %(i+1), ".tab")

    new_report = cir.post.reports_by_category.spectral("ang_rad(V(Port%d_T2))" %(i+1))
    new_report.window = "Hanning"
    new_report.max_frequency = "100MHz"
    new_report.time_start = "1us"
    new_report.time_stop = "11us"
    new_report.create()
    new_report.edit_x_axis_scaling(units='Hz')
    new_report.edit_x_axis_scaling(linear_scaling=False)
    cir.post.rename_report(new_report.plot_name, "Port%d_T2_ang" %(i+1))
    cir.post.export_report_to_file(wd, "Port%d_T2_ang" %(i+1), ".tab")

cir.save_project()
cir.release_desktop(close_projects=False, close_desktop=False)

# Reset sources in HFSS
for i in range(int(npins/2)):
    h.edit_source(portandmode="Port%d_T1" %(i+1), powerin="0V", phase="0deg")
    h.edit_source(portandmode="Port%d_T2" %(i+1), powerin="0V", phase="0deg")

h.save_project()

# Delete dataset
for key in h.design_datasets.keys():
    h.odesign.DeleteDataset(key)

# Import all ".tab" files in the Dataset
for file in os.listdir(wd):
    if file.endswith(".tab"):
        file = os.path.join(wd, file)
        name = os.path.splitext(os.path.basename(file))[0]
        h.import_dataset1d(file, dsname=name, is_project_dataset=False)

# Create excitations using Datasets
for i in range(int(npins/2)):
    h.edit_source(portandmode="Port%d_T1" %(i+1), powerin="pwl(Port%d_T1_mag,Freq)" %(i+1), phase="pwl(Port%d_T1_ang,Freq)" %(i+1))
    h.edit_source(portandmode="Port%d_T2" %(i+1), powerin="pwl(Port%d_T2_mag,Freq)" %(i+1), phase="pwl(Port%d_T2_ang,Freq)" %(i+1))

h.analyze_setup("Setup1")
h.save_project()
h.release_desktop(close_projects=False, close_desktop=False)
