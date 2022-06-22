import sys
import os
import time
from mygui import *
from Readout_Sheets_HFSS_Widgets import mySheetsQTableWidget, myStandardItemModel, myFormWidget
from PyQt5 import QtWidgets, QtCore, QtGui
import ctypes
co_initialize = ctypes.windll.ole32.CoInitialize
co_initialize(None)
import pyaedt
import pyvistaqt as pvqt
# import openpyxl

sys.setrecursionlimit(100_000)
os.environ["QT_API"] = "pyqt5"
print(chr(27) + "[2J")

class DesignerMainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        # Initialization of the superclass
        super(DesignerMainWindow, self).__init__(parent)
        self.setupUi(self)

        # Define Variable
        self.filepath = ''
        self.nongraphical = False
        self.hfssversion = ''
        self.h = 0
        self.plot_obj = 0
        self.sel_proj_model_item = ''
        self.sel_cond_model_item = ''
        self.sel_sheet_item = ''

        self.readDesignPushButton.hide()
        self.graphCheckBox.setEnabled(False)

        # Creating models
        self.project_model = myStandardItemModel()

        # Creating a default empty 'sheets' QTableWidget
        self.sheetsTableWidget = QtWidgets.QTableWidget()
        self.gridLayout_3.addWidget(self.sheetsTableWidget)

        # Creating 'design' QTableView
        self.projectView = QtWidgets.QListView()
        self.projectView.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout_1.addWidget(self.projectView)
        self.projectView.setModel(self.project_model)

        # Creating Plotter
        self.plotter = pvqt.QtInteractor(self.frame)
        layout3 = QtWidgets.QVBoxLayout()
        layout3.addWidget(self.plotter.interactor)
        self.frame.setLayout(layout3)

        # Signals / Slots connections
        self.actionAedtLoad.triggered.connect(self.loadAedt)
        self.actionAedtClose.triggered.connect(self.closeAedt)
        self.actionAedtSave.triggered.connect(self.saveAedt)
        self.project_model.myModelSignal.connect(self.getProjModelItem)
        self.graphCheckBox.stateChanged.connect(self.plotPreview)

    def isQuitClicked(self):
        sys.exit(app.exec_())

    def closeEvent(self, event): # Use if the main window is closed by the user
        close = QtWidgets.QMessageBox.question(self, "QUIT", "Confirm quit?", QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if close == QtWidgets.QMessageBox.Yes:
            self.closeAedt()
            event.accept()
            app.quit()
        else:
            event.ignore()

    def loadAedt(self):
        self.graphModeCheckBox.setEnabled(False)
        self.filepath, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Load \".aedt\" File", "", "Config file (*.aedt)")
        if os.path.isfile(self.filepath):
            self.nongraphical = self.graphModeCheckBox.isChecked()
            self.hfssversion = self.hfssVerComboBox.currentText()
            self.readDesignPushButton.setEnabled(False)
            pyaedt.generic.general_methods.remove_project_lock(self.filepath)

            if self.hfssversion != '':
                self.statusBar.showMessage("Loading AEDT project. Please wait...")
                QtCore.QCoreApplication.processEvents()
                # try:
                # self.h = pyaedt.Hfss()
                # self.h.set_active_design("Model_LumpedPorts_1")
                self.h = pyaedt.Hfss(projectname=self.filepath, specified_version=self.hfssversion, non_graphical=self.nongraphical, new_desktop_session=True)
                self.sheetsTableWidget.deleteLater()  # delete Table widget
                self.sheetsTableWidget = mySheetsQTableWidget(self.h)
                self.gridLayout_3.addWidget(self.sheetsTableWidget)

                self.populateDesignlist()
                self.statusBar.showMessage("Project Loaded.")

                # QtWidgets.QMessageBox.about(self, "STATUS", "Project loaded correctly.")
                self.readDesignPushButton.setEnabled(True)
                self.actionAedtLoad.setEnabled(False)
                # except:
                #     self.statusBar.showMessage("Project was not loaded... Did you choose the right HFSS version?")
                #     QtWidgets.QMessageBox.about(self, "STATUS", "Project was not loaded... Did you choose the right HFSS version?")
        else:
            print("No File selected")

    def populateDesignlist(self):
        designs = self.h.design_list
        for i in range(len(designs)):
            item = QtGui.QStandardItem()
            item.setText(designs[i])
            item.setFlags(QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
            item.setData(QtCore.Qt.Unchecked, QtCore.Qt.CheckStateRole)
            self.project_model.appendRow(item)
        self.graphCheckBox.setEnabled(True)
        self.project_model.item(0,0).setCheckState(QtCore.Qt.Checked) # activate first design from the project_list
        self.h.set_active_design(designs[0])
        self.readDesign()

    def plotPreview(self):
        if self.graphCheckBox.isChecked() == True:
            if self.plot_obj == 0:
                self.plot_obj = self.h.plot(show=False, plot_as_separate_objects=True)
                self.plot_obj.plot(r"c:\temp\img.jpg")
            for objs in self.plot_obj.objects:
                color_cad = [i / 255 for i in objs.color]
                self.plotter.add_mesh(objs._cached_polydata, color=color_cad, opacity=objs.opacity)
        else :
            self.plotter.clear()

    def readDesign(self):
        self.sheetsTableWidget.setRowCount(0)
        self.h.set_active_design(self.sel_proj_model_item) # self.h.set_active_design(self.h.design_list[0])
        self.sheetsTableWidget.fillSheetsTable()
        self.plotPreview()

    def getProjModelItem(self, sel_proj_model_item):
        self.sel_proj_model_item = sel_proj_model_item
        self.h.save_project()
        self.plotter.clear()
        self.plot_obj = 0
        self.readDesign()

    def saveAedt(self):
        self.h.save_project()

    def closeAedt(self):
        if os.name != "posix":
            if  self.h != 0:
                self.h.release_desktop()
            self.actionAedtLoad.setEnabled(True)
            self.sheetsTableWidget.setRowCount(0)
            self.project_model.removeRows(0, self.project_model.rowCount())
            self.plotter.clear()
            self.plot_obj = 0
            self.graphCheckBox.setChecked(False)
            self.graphCheckBox.setEnabled(False)
            self.graphModeCheckBox.setEnabled(True)


########################################################################################################################

if __name__ == '__main__':
    if not QtWidgets.QApplication.instance():
        app = QtWidgets.QApplication(sys.argv)
    else:
        app = QtWidgets.QApplication.instance()
    # app.setStyle('Fusion')  # 'Breeze', 'Oxygen', 'QtCurve', 'Windows', 'Fusion'
    app.setStyle('Fusion')
    w = DesignerMainWindow()
    w.show()
    sys.exit(app.exec_())
