import os
import numpy as np
from PyQt5 import QtWidgets, QtCore, QtGui
import matplotlib
matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
from matplotlib.figure import Figure


#######################################################################################################################
class MplCanvas(FigureCanvasQTAgg):
    def __init__(self, parent=None, width=5, height=4, dpi=30):
        fig = Figure(figsize=(width, height), dpi=dpi)
        self.ax = fig.add_subplot(111)
        self.ax.axison = False
        super(MplCanvas, self).__init__(fig)

#######################################################################################################################
class myStandardItemModel(QtGui.QStandardItemModel):
    myModelSignal = QtCore.pyqtSignal(object)
    def setData(self, index, value, role=QtCore.Qt.EditRole):
        result = super(myStandardItemModel, self).setData(index, value, role)
        cur_model_val = index.data(value)
        self.myModelSignal.emit(cur_model_val)
        for i in range(self.rowCount()): # only one element can be checked
            if self.item(i).data(value) != cur_model_val:
                self.item(i).setCheckState(QtCore.Qt.Unchecked)
        return True

#######################################################################################################################
class myComponentsQTableWidget(QtWidgets.QTableWidget):
    def __init__(self, parent=None):
        super(myComponentsQTableWidget, self).__init__(parent)
        labels = ["Name", "Layer", "Position", "Angle (deg)", "Pin number", "Replace Component"]
        self.setAlternatingRowColors(True)
        self.setRowCount(100)
        self.setColumnCount(len(labels))
        self.horizontalHeader().setStretchLastSection(True)
        # self.verticalHeader().setStretchLastSection(True)
        self.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.setHorizontalHeaderLabels(labels)
        self.h3d = 0
        self.edb = 0
        # self.cond_model = myStandardItemModel()
        # self.itemChanged.connect(self.renameItem)

    def setH3d(self, h3d):
        self.h3d = h3d
        self.form = myFormWidget(self.h3d)
        self.form.myFormSignal.connect(self.import3DComponent)

    def setEdb(self, edb):
        self.edb = edb

    def fillComponentsTable(self):
        self.setRowCount(0)
        i = 0
        objects = self.edb.core_components.components
        for obj_id, o_list in objects.items():  # fill in Sheets QtableWidget
            self.insertRow(self.rowCount())
            s_array = [obj_id, o_list.placement_layer, o_list.center, o_list.edbcomponent.GetTransform().Rotation.ToDouble()*180/np.pi, o_list.numpins] # Name", "Layer", "Position", "Angle"
            for j in range(len(s_array)):
                item = QtWidgets.QTableWidgetItem(str(s_array[j]))
                item.setFlags(QtCore.Qt.ItemIsEnabled)
                # if j != 3:
                #     item.setFlags(QtCore.Qt.ItemIsEnabled)
                self.setItem(i, j, item)
            i = i + 1

    def contextMenuEvent(self, event):
        if self.currentColumn() == 5:
            contextMenu = QtWidgets.QMenu(self)
            Act1 = contextMenu.addAction("Place 3D Component")
            Act2 = contextMenu.addAction("Remove Component")
            action = contextMenu.exec_(self.mapToGlobal(event.pos()))
            if action == Act1:
                if self.h3d != 0 :
                    self.form.show()
                else :
                    msg = QtWidgets.QMessageBox()
                    msg.setIcon(QtWidgets.QMessageBox.Warning)
                    msg.setText("Please export the '.def' file to 'HFSS3DLayout' before adding components.")
                    msg.exec()

            if action == Act2:
                self.deleteComponent()

    def import3DComponent(self, comp_path, comp_angle):
        layer = self.item(self.currentRow(), 1).text()
        comp_design_name = self.item(self.currentRow(), 0).text()
        comp3d_name = os.path.splitext(os.path.basename(comp_path))[0]
        comp_name = self.item(self.currentRow(), 0).text()

        # set pin number in the QTableWidget
        item = QtWidgets.QTableWidgetItem()
        numpins = self.item(self.currentRow(), 4).text()

        # separate pos string to float x,y
        pos = self.item(self.currentRow(), 2).text()
        for character in "[]'":
            pos = pos.replace(character, "")
        pos = pos.split(',')

        # Place a 3d component
        comp3d = self.h3d.modeler.place_3d_component(comp_path, number_of_terminals=numpins, placement_layer=layer, component_name=comp3d_name, pos_x=pos[0], pos_y=pos[1])
        comp3d.angle = "%.1f" %float(comp_angle)

        # set new angle in the QTableWidget
        item = QtWidgets.QTableWidgetItem(comp_angle)
        self.setItem(self.currentRow(), 3, item)

        # set component name in the QTableWidget
        item = QtWidgets.QTableWidgetItem(comp3d_name)
        self.setItem(self.currentRow(), 5, item)

        # set component name in the QTableWidget
        item = QtWidgets.QTableWidgetItem(comp3d.name)
        self.setItem(self.currentRow(), 0, item)

    def deleteComponent(self):
        comp_name = self.item(self.currentRow(), 0).text()
        self.h3d.modeler.oeditor.DissolveComponents(["NAME:elements", comp_name])
        item = QtWidgets.QTableWidgetItem("")
        self.setItem(self.currentRow(), 5, item)


#######################################################################################################################
class myFormWidget(QtWidgets.QWidget):
    myFormSignal = QtCore.pyqtSignal(object, object)
    def __init__(self, h3d, parent=None):
        super(myFormWidget, self).__init__(parent)
        self.setObjectName("Form")
        self.resize(380, 125)
        self.setWindowTitle("")
        self.gridLayout = QtWidgets.QGridLayout(self)
        self.gridLayout.setObjectName("gridLayout")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.label_2 = QtWidgets.QLabel(self)
        self.label_2.setObjectName("label_2")
        self.label_2.setText("Choose '.a3dcomp' file:")
        self.verticalLayout.addWidget(self.label_2)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.BrowsePushButton = QtWidgets.QPushButton(self)
        self.BrowsePushButton.setObjectName("BrowsePushButton")
        self.BrowsePushButton.setText("Browse")
        self.horizontalLayout_3.addWidget(self.BrowsePushButton)
        self.CompChoiceLineEdit = QtWidgets.QLineEdit(self)
        self.CompChoiceLineEdit.setObjectName("CompChoiceLineEdit")
        self.horizontalLayout_3.addWidget(self.CompChoiceLineEdit)
        self.verticalLayout.addLayout(self.horizontalLayout_3)
        self.verticalLayout_2.addLayout(self.verticalLayout)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(self)
        self.label.setObjectName("label")
        self.label.setText("Rotation angle (deg)")
        self.horizontalLayout.addWidget(self.label)
        self.RotationLineEdit = QtWidgets.QLineEdit(self)
        self.RotationLineEdit.setObjectName("RotationLineEdit")
        self.RotationLineEdit.setText("0")
        self.horizontalLayout.addWidget(self.RotationLineEdit)
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem)
        self.verticalLayout_2.addLayout(self.horizontalLayout)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.SubmitPushButton = QtWidgets.QPushButton(self)
        self.SubmitPushButton.setObjectName("SubmitPushButton")
        self.SubmitPushButton.setText("Submit")
        self.horizontalLayout_2.addWidget(self.SubmitPushButton)
        self.CancelPushButton = QtWidgets.QPushButton(self)
        self.CancelPushButton.setObjectName("CancelPushButton")
        self.CancelPushButton.setText("Cancel")
        self.horizontalLayout_2.addWidget(self.CancelPushButton)
        self.verticalLayout_2.addLayout(self.horizontalLayout_2)
        self.gridLayout.addLayout(self.verticalLayout_2, 0, 0, 1, 1)

        self.SubmitPushButton.clicked.connect(self.onSubmitClicked)
        self.CancelPushButton.clicked.connect(self.onCancelClicked)
        self.BrowsePushButton.clicked.connect(self.onBrowseClicked)
        # self.cond_model.myModelSignal.connect(self.getSignalFromMyStandardItemModel)
        self.comp3d_path = ''
        self.edb = 0

    def onBrowseClicked(self):
        self.comp3d_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Load \"a3dcomp\" file", "", "3D Component file (*.a3dcomp)")
        self.CompChoiceLineEdit.setText(self.comp3d_path)

    def onSubmitClicked(self):
        self.myFormSignal.emit(self.comp3d_path, self.RotationLineEdit.text())
        self.close()

    def onCancelClicked(self):
        self.close()

    def on_close(self):
        self.close()

#######################################################################################################################