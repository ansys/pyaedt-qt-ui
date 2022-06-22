from PyQt5 import QtWidgets, QtCore, QtGui

def createLumpedPort(h, port_name, sheet_item, conductor_item):
    h.create_lumped_port_to_sheet(sheet_item, axisdir=0, impedance=50, portname=port_name, renorm=True, deemb=False, reference_object_list=[conductor_item])

def deletePort(h, object_name):
    for i in range(len(h.boundaries)):
        if "Faces" in h.boundaries[i].props:
            port_id = h.boundaries[i].props["Faces"][0]
            port_type = h.boundaries[i].type
            port_objectname = h.modeler.oeditor.GetObjectNameByFaceID(port_id)
            if port_objectname == object_name:
                h.boundaries[i].delete()
                break
    return 0

def getConductorList(h, object_name): # get all conductors in contact with a given sheet
    # cond_list = getAllConductors(h)
    vertexes = h.modeler.get_object_vertices(object_name)
    obj = []
    units = h.modeler.model_units
    h.modeler.set_working_coordinate_system('Global')
    for i in range(4):
        a = h.modeler.get_vertex_position(vertexes[i])
        b = h.modeler.get_bodynames_from_position(a, units)
        for j in range(len(b)):
            obj.append(b[j])
            # if h.modeler.get_object_from_name(b[j])._m_name in cond_list:
            #     obj.append(b[j])
        if object_name in obj:
            obj.remove(object_name)
    return(list(set(obj)))

def getSheetsStatus(h):
    port_dict = {}
    sheets_list = h.modeler.sheet_names
    for i in range(len(sheets_list)):
        port_dict[sheets_list[i]] = "Unassigned"

    for i in range(len(h.boundaries)):
        if "Faces" in h.boundaries[i].props:
            port_id = h.boundaries[i].props["Faces"][0]
            port_type = h.boundaries[i].type
            port_objectname = h.modeler.oeditor.GetObjectNameByFaceID(port_id)
            port_dict[port_objectname] = port_type
        if "Objects" in h.boundaries[i].props:
            port_id = h.boundaries[i].props["Objects"][0]
            port_type = h.boundaries[i].type
            port_objectname = h.modeler.oeditor.GetObjectNameByID(port_id)
            port_dict[port_objectname] = port_type
    return port_dict

# def getAllConductors(h): # get all project conductors & conducting boundaries
#     cond_list = h.modeler._materials.conductors
#     cond_dict = {}
#     objects = h.modeler.objects
#     for obj_id, o_list in objects.items():
#         if o_list._material_name in cond_list:
#             cond_dict[obj_id] = o_list._m_name
#     for i in range(len(h.boundaries)):
#         if "Perfect E" in h.boundaries[i].props["BoundType"] or "Finite Conductivity" in h.boundaries[i].props["BoundType"]:
#             if "Faces" in h.boundaries[i].props:
#                 cond_ID = h.boundaries[i].props["Faces"][0]
#                 cond_objectname = h.modeler.oeditor.GetObjectNameByFaceID(cond_ID)
#             elif "Objects" in h.boundaries[i].props:
#                 cond_ID = h.boundaries[i].props["Objects"][0]
#                 cond_objectname = h.modeler.oeditor.GetObjectNameByID(cond_ID)
#             cond_dict[cond_ID] = cond_objectname
#     return list(cond_dict.values())

#######################################################################################################################

class myStandardItemModel(QtGui.QStandardItemModel):
    myModelSignal = QtCore.pyqtSignal(object)
    def setData(self, index, value, role=QtCore.Qt.EditRole):
        result = super(myStandardItemModel, self).setData(index, value, role)
        cur_model_val = index.data(value)
        self.myModelSignal.emit(cur_model_val)
        for i in range(self.rowCount()): # Only one element can be checked
            if self.item(i).data(value) != cur_model_val:
                self.item(i).setCheckState(QtCore.Qt.Unchecked)
        return True

#######################################################################################################################

class mySheetsQTableWidget(QtWidgets.QTableWidget):
    def __init__(self, h, parent=None):
        super(mySheetsQTableWidget, self).__init__(parent)
        labels = ["ID", "Name", "Color", "Object Type", "Sheet Status"] # 'lumped port', 'not assigned' oder 'Conductor' (PEC, Finite Conductivity)
        self.setAlternatingRowColors(True)
        self.setRowCount(100)
        self.setColumnCount(len(labels))
        self.horizontalHeader().setStretchLastSection(True)
        # self.verticalHeader().setStretchLastSection(True)
        self.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.setHorizontalHeaderLabels(labels)
        self.h = h
        self.port_name = ''
        self.cond_model = myStandardItemModel()
        self.form = myFormWidget(self.h)
        self.form.enterButton.clicked.connect(self.onEnterClicked)
        self.form.myFormSignal.connect(self.getSignalfromForm)
        self.itemChanged.connect(self.renameItem)

    def getSignalfromForm(self, obj1, obj2):
        self.port_name = obj1
        self.sel_cond_item = obj2

    def contextMenuEvent(self, event):
        contextMenu = QtWidgets.QMenu(self)
        Act1 = contextMenu.addAction("Create Lumped Port from sheet")
        Act2 = contextMenu.addAction("Delete Lumped Port")
        action = contextMenu.exec_(self.mapToGlobal(event.pos()))
        if action == Act1:
            self.sel_sheet_item = self.item(self.currentRow(), 1).text()
            self.form.createRefCondList(self.sel_sheet_item)
            self.form.setWindowModality(QtCore.Qt.ApplicationModal)
            self.form.show()
        if action == Act2:
            port_dict = getSheetsStatus(self.h)
            self.sel_sheet_item = self.item(self.currentRow(), 1).text()
            key = list(port_dict.keys()).index(self.sel_sheet_item)
            print("Delete port: %s, %s" % (key, self.sel_sheet_item))
            deletePort(self.h, self.sel_sheet_item)
            self.fillSheetsTable()

    def fillSheetsTable(self):
        self.itemChanged.disconnect()
        self.setRowCount(0)
        port_dict = getSheetsStatus(self.h)
        i = 0
        objects = self.h.modeler.objects
        for obj_id, o_list in objects.items():  # fill in Sheets QtableWidget
            assigned_to_port = ''
            if o_list._m_name in self.h.modeler._sheets:
                self.insertRow(self.rowCount())
                key = list(port_dict.keys()).index(o_list._m_name)
                assigned_to_port = port_dict[o_list._m_name]
                s_array = [obj_id, o_list._m_name, o_list._color, "Sheet", assigned_to_port]
                for j in range(len(s_array)):
                    item = QtWidgets.QTableWidgetItem(str(s_array[j]))
                    if j != 1: # make all column not editable except column "name"
                        item.setFlags(QtCore.Qt.ItemIsEnabled)
                    if j == 2: # change background color of the "color" column
                        item.setBackground(QtGui.QColor(o_list.color[0], o_list.color[1], o_list.color[2]))
                    self.setItem(i, j, item)
                i = i + 1
        self.itemChanged.connect(self.renameItem)

    def onEnterClicked(self):
        self.sel_sheet_item = self.item(self.currentRow(), 1).text()
        createLumpedPort(self.h, self.port_name, self.sel_sheet_item, self.sel_cond_item)
        self.fillSheetsTable()

    def renameItem(self):
        if self.currentColumn() == 1: # only the 'name' column can be renamed
            id = int(self.item(self.currentRow(), 0).text())
            val = self.item(self.currentRow(), 1).text()
            self.h.modeler.objects[id].name = val

#######################################################################################################################

class myFormWidget(QtWidgets.QWidget):
    myFormSignal = QtCore.pyqtSignal(object, object)
    def __init__(self, h, parent=None):
        super(myFormWidget, self).__init__(parent)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.setFont(font)
        self.setObjectName("Form")
        # self.setWindowFlag(QtCore.Qt.FramelessWindowHint)
        self.setWindowTitle('Lumped port settings')
        self.resize(497, 552)
        self.cond_model = myStandardItemModel()
        self.gridLayout = QtWidgets.QGridLayout(self)
        self.gridLayout.setObjectName("gridLayout")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(self)
        self.label.setObjectName("label")
        self.label.setText("Port Name:")
        self.horizontalLayout.addWidget(self.label)
        self.lineEdit = QtWidgets.QLineEdit(self)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout.addWidget(self.lineEdit)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.label_2 = QtWidgets.QLabel(self)
        self.label_2.setObjectName("label_2")
        self.label_2.setText("Choose the reference conductor:")
        self.verticalLayout.addWidget(self.label_2)
        self.listView = QtWidgets.QListView(self)
        self.listView.setObjectName("listWidget")
        self.verticalLayout.addWidget(self.listView)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.enterButton = QtWidgets.QPushButton(self)
        self.enterButton.setObjectName("Enter")
        self.enterButton.setText("Enter")
        self.horizontalLayout_2.addWidget(self.enterButton)
        self.cancelButton = QtWidgets.QPushButton(self)
        self.cancelButton.setObjectName("Cancel")
        self.cancelButton.setText("Cancel")
        self.horizontalLayout_2.addWidget(self.cancelButton)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.gridLayout.addLayout(self.verticalLayout, 0, 0, 1, 1)

        self.enterButton.clicked.connect(self.onEnterClicked)
        self.cancelButton.clicked.connect(self.onCancelClicked)
        self.cond_model.myModelSignal.connect(self.getSignalFromMyStandardItemModel)
        self.h = h

    def getSignalFromMyStandardItemModel(self, item):
        self.cur_model_val = item

    def createRefCondList(self, sheet_item):
        self.sel_sheet_item = sheet_item
        self.cond_model.removeRows(0, self.cond_model.rowCount())
        self.obj_cond_list = []
        self.obj_cond_list = list(getConductorList(self.h, self.sel_sheet_item))

        for i in range(len(self.obj_cond_list)):  # create Conductor model
            item = QtGui.QStandardItem()
            item.setText(str(self.obj_cond_list[i]))
            item.setFlags(QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
            item.setData(QtCore.Qt.Unchecked, QtCore.Qt.CheckStateRole)
            self.cond_model.appendRow(item)
        self.listView.setModel(self.cond_model)

    def onEnterClicked(self):
        self.myFormSignal.emit(self.lineEdit.text(), self.cur_model_val)
        self.close()

    def onCancelClicked(self):
        self.close()

    def on_close(self):
        self.close()

#######################################################################################################################

