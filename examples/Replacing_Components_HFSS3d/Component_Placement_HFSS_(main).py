# This program will allow the user to do the following :
# 1- Load an odb++ file and create a ".def" file,
# 2- To preview the odb++ file and to generate a component list (position, name etc...)
# 3- Create a HFSS3D project from the ".def" file
# 4- Once the HFSS3D project is created then the user can delete components or place 3D components (Right click the last column "Replace Component")

import sys
import os
import time
from Component_Placement_HFSS_GUI import *
from Component_Placement_HFSS_Widgets import myComponentsQTableWidget, myStandardItemModel, MplCanvas
from PyQt5 import QtWidgets, QtCore, QtGui
import ctypes
co_initialize = ctypes.windll.ole32.CoInitialize
co_initialize(None)
import pyaedt
from pyaedt.generic.general_methods import convert_remote_object
import matplotlib.pyplot as plt
from matplotlib.path import Path
from matplotlib.patches import PathPatch


sys.setrecursionlimit(100_000)
os.environ["QT_API"] = "pyqt5"
print(chr(27) + "[2J")

class DesignerMainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        # Initialization of the superclass
        super(DesignerMainWindow, self).__init__(parent)
        self.setupUi(self)

        # Define Variable
        self.odb_path = ''
        self.working_directory = ''
        self.nongraphical = False
        self.hfssversion = ''
        self.edb = 0
        self.h3d = 0

        self.readDesignPushButton.hide()
        # self.graphCheckBox.setEnabled(False)
        self.graphCheckBox.setChecked(False)

        self.actionOdbLoad.setEnabled(True)
        self.actionEdbClose.setEnabled(False)
        self.actionAedtCreate.setEnabled(False)
        self.actionHFSS3DLayoutClose.setEnabled(False)

        # Creating models
        self.project_model = myStandardItemModel()

        # Creating a default empty 'sheets' QTableWidget
        self.componentsTableWidget = QtWidgets.QTableWidget()
        self.gridLayout_3.addWidget(self.componentsTableWidget)

        # Creating 'design' QTableView
        self.projectView = QtWidgets.QListView()
        self.projectView.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout_1.addWidget(self.projectView)
        self.projectView.setModel(self.project_model)

        # Creating Plotter
        self.plotter = MplCanvas(self, width=5, height=4, dpi=100)
        layout3 = QtWidgets.QVBoxLayout()
        layout3.addWidget(self.plotter)
        self.frame.setLayout(layout3)

        # Signals / Slots connections
        self.actionOdbLoad.triggered.connect(self.loadOdb)
        # self.actionEdbClose.triggered.connect(self.closeEdb)
        self.actionAedtCreate.triggered.connect(self.createHFSS3dLayout)
        self.actionHFSS3DLayoutClose.triggered.connect(self.close)
        # self.project_model.myModelSignal.connect(self.getProjModelItem)
        self.graphCheckBox.stateChanged.connect(self.plotPreview)

    def isQuitClicked(self):
        sys.exit(app.exec_())

    def closeEvent(self, event): # Use if the main window is closed by the user
        close = QtWidgets.QMessageBox.question(self, "QUIT", "Confirm quit?", QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if close == QtWidgets.QMessageBox.Yes:
            self.close()
            event.accept()
            app.quit()
        else:
            event.ignore()

    def loadOdb(self):
        self.graphModeCheckBox.setEnabled(False)
        self.odb_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Load \"odb++\" File", "", "CAD file (*.tgz)")
        self.odb_name = os.path.splitext(os.path.basename(self.odb_path))[0]

        if os.path.isfile(self.odb_path):
            self.nongraphical = self.graphModeCheckBox.isChecked()
            self.hfssversion = self.hfssVerComboBox.currentText()
            self.readDesignPushButton.setEnabled(False)

            self.working_directory = os.getcwd()
            pyaedt.generic.general_methods.remove_project_lock(self.working_directory)

            if self.hfssversion != '':
                self.statusBar.showMessage("Loading ODB++ file. Please wait...")
                QtCore.QCoreApplication.processEvents()
                self.edb = pyaedt.Edb(edbpath=self.working_directory, edbversion=self.hfssversion)
                self.edb.import_layout_pcb(self.odb_path, self.working_directory)

                self.componentsTableWidget.deleteLater()  # delete Table widget
                self.componentsTableWidget = myComponentsQTableWidget()
                self.componentsTableWidget.setEdb(self.edb)
                self.gridLayout_3.addWidget(self.componentsTableWidget)

                self.populateDesignlist()
                self.statusBar.showMessage("ODB++ file loaded.")

                self.readDesignPushButton.setEnabled(True)
                self.actionOdbLoad.setEnabled(False)
                self.actionEdbClose.setEnabled(True)
                self.actionAedtCreate.setEnabled(True)
                self.actionHFSS3DLayoutClose.setEnabled(False)
        else:
            print("No File selected")

    def populateDesignlist(self):
        cells = self.edb.cell_names
        for i in range(len(cells)):
            item = QtGui.QStandardItem()
            item.setText(cells[i])
            self.project_model.appendRow(item)
        self.graphCheckBox.setEnabled(True)
        self.readDesign()

    def plotPreview(self):
        if self.graphCheckBox.isChecked() == True:
            start = time.time()
            object_lists = self.edb.core_nets.get_plot_data(None)
            self.plot_matplotlib(object_lists)
            duration = time.time() - start
            print(duration)
        else:
            self.plotter.ax.clear()
            self.plotter.ax.axison = False
            self.plotter.draw()

    def readDesign(self):
        self.componentsTableWidget.setRowCount(0)
        self.componentsTableWidget.fillComponentsTable()
        self.plotPreview()

    def createHFSS3dLayout(self):
        if self.edb != 0:
            def_file_name, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Load \"def\" file", "", "CAD file (*.def)")
            if os.path.isfile(def_file_name):
                self.statusBar.showMessage("Creating the HFSS3DLayout project. Please wait...")
                app.processEvents()
                # delete ".aedt" and "aedt.lock" files
                pyaedt.generic.general_methods.remove_project_lock(self.working_directory)
                aedt_file_name = self.working_directory + "\\" + self.odb_name + ".aedt"
                if os.path.exists(aedt_file_name):
                    os.remove(aedt_file_name)
                # save and close the edb file
                self.edb.save_edb()
                self.edb.close_edb()
                self.edb = 0
                # create the HFSS3DLayout project
                self.h3d = pyaedt.Hfss3dLayout(projectname=def_file_name, specified_version=self.hfssversion, non_graphical=self.nongraphical, new_desktop_session=True)
                self.componentsTableWidget.setH3d(self.h3d)
                # update status bar
                self.statusBar.showMessage("HFSS3DLayout project created.")
                app.processEvents()
                # # save and close the edb file
                # self.edb.save_edb()
                # self.edb.close_edb()
                # self.edb = 0
                # update GUI
                self.graphCheckBox.setDisabled(True)
                self.actionOdbLoad.setEnabled(False)
                self.actionEdbClose.setEnabled(False)
                self.actionAedtCreate.setEnabled(False)
                self.actionHFSS3DLayoutClose.setEnabled(True)

    def close(self):
        if os.name != "posix":
            if self.edb != 0:
                self.edb.save_edb()
                self.edb.close_edb()

            if self.h3d != 0:
                self.graphCheckBox.setChecked(False)
                self.graphModeCheckBox.setEnabled(True)
                self.h3d.close_project()
                self.h3d.release_desktop()

        self.actionOdbLoad.setEnabled(True)
        # self.actionEdbClose.setEnabled(False)
        self.actionAedtCreate.setEnabled(False)
        self.actionHFSS3DLayoutClose.setEnabled(False)
        self.componentsTableWidget.setRowCount(0)
        self.project_model.removeRows(0, self.project_model.rowCount())
        self.plotter.ax.clear()
        self.plotter.ax.axison = False
        self.plotter.draw()
        self.graphCheckBox.setDisabled(False)
        self.graphCheckBox.setChecked(False)
        self.graphModeCheckBox.setEnabled(True)


    # def closeEdb(self):
    #     if os.name != "posix":
    #         if self.edb != 0:
    #             self.edb.save_edb()
    #             self.edb.close_edb()
    #         self.actionOdbLoad.setEnabled(True)
    #         self.actionEdbClose.setEnabled(False)
    #         self.actionAedtCreate.setEnabled(False)
    #         self.actionHFSS3DLayoutClose.setEnabled(False)
    #         self.componentsTableWidget.setRowCount(0)
    #         self.project_model.removeRows(0, self.project_model.rowCount())
    #         self.plotter.ax.clear()
    #         self.plotter.ax.axison = False
    #         self.plotter.draw()
    #         self.graphCheckBox.setChecked(False)
    #         self.graphModeCheckBox.setEnabled(True)

    def plot_matplotlib(self, plot_data, size=(100, 100), show_legend=True, xlabel="", ylabel="", title="", snapshot_path=None):
        self.plotter.ax.clear()
        self.plotter.ax.axison = False
        self.plotter.draw()
        try:
            len(plot_data)
        except:
            plot_data = convert_remote_object(plot_data)
        for object in plot_data:
            if object[-1] == "fill":
                self.plotter.ax.fill(object[0], object[1], c=object[2], label=object[3], alpha=object[4])
            elif object[-1] == "path":
                path = Path(object[0], object[1])
                patch = PathPatch(path, color=object[2], alpha=object[4], label=object[3])
                self.plotter.ax.add_patch(patch)

        self.plotter.ax.set(xlabel=xlabel, ylabel=ylabel, title=title)
        if show_legend:
            self.plotter.ax.legend()
        self.plotter.ax.axis("equal")
        self.plotter.ax.axison = False
        self.plotter.draw()

########################################################################################################################

if __name__ == '__main__':
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    if not QtWidgets.QApplication.instance():
        app = QtWidgets.QApplication(sys.argv)
    else:
        app = QtWidgets.QApplication.instance()
    app.setStyle('Fusion')  # 'Breeze', 'Oxygen', 'QtCurve', 'Windows', 'Fusion'
    w = DesignerMainWindow()
    w.show()
    sys.exit(app.exec_())
