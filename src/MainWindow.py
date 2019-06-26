# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'MainWindow.ui'
#
# Created by: PyQt5 UI code generator 5.10.1
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import *
import xlrd


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(880, 600)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(30, 20, 61, 31))
        self.label.setObjectName("label")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(110, 30, 531, 21))
        self.lineEdit.setObjectName("lineEdit")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(690, 30, 60, 20))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.openfile)

        self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget.setGeometry(QtCore.QRect(30, 90, 621, 431))
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.tableWidget.itemClicked.connect(self.get_sum)
        self.tableWidget.horizontalHeader().sectionClicked.connect(self.table_hidden_changed)
        self.tableWidget.verticalHeader().sectionClicked.connect(self.table_vertical_state_changed)
        self.last_check_state = []

        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(690, 470, 150, 30))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(self.save_result)

        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(690, 400, 60, 20))
        self.label_2.setObjectName("label_2")
        self.comboBox3 = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox3.setGeometry(QtCore.QRect(770, 400, 80, 20))
        self.comboBox3.setObjectName("comboBox")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(690, 430, 60, 20))
        self.label_3.setObjectName("label_3")
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(770, 430, 80, 20))
        self.label_6.setObjectName("label_6")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(690, 160, 60, 20))
        self.label_4.setObjectName("label_4")
        self.comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox.setGeometry(QtCore.QRect(690, 190, 150, 30))
        self.comboBox.setObjectName("comboBox")
        self.comboBox.activated.connect(lambda: self.creat_table_show(self.comboBox.currentIndex()))
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(690, 240, 60, 20))
        self.label_5.setObjectName("label_5")
        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setGeometry(QtCore.QRect(760, 240, 80, 20))
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_4.clicked.connect(self.table_show_colnums)
        self.pushButton_5 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_5.setGeometry(QtCore.QRect(760, 320, 80, 20))
        self.pushButton_5.setObjectName("pushButton_5")
        self.pushButton_5.clicked.connect(self.table_show_verticals)
        self.comboBox2 = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox2.setGeometry(QtCore.QRect(690, 270, 150, 30))
        self.comboBox2.setObjectName("comboBox2")
        self.comboBox2.activated.connect(lambda: self.table_hidden_changed(self.comboBox2.currentIndex()))
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(690, 90, 150, 60))
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_3.clicked.connect(lambda: self.creat_table_show(0))
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 787, 23))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "研发配盘"))
        self.label.setText(_translate("MainWindow", "原始文件"))
        self.pushButton.setText(_translate("MainWindow", "选择文件"))
        self.pushButton_2.setText(_translate("MainWindow", "结果保存"))
        self.label_2.setText(_translate("MainWindow", "金额所在列"))
        self.label_3.setText(_translate("MainWindow", "总金额"))
        self.label_6.setText(_translate("MainWindow", "0.0"))
        self.label_4.setText(_translate("MainWindow", "Sheet选择"))
        self.label_5.setText(_translate("MainWindow", "修改列状态"))
        self.pushButton_3.setText(_translate("MainWindow", "预览"))
        self.pushButton_4.setText(_translate("MainWindow", "显示所有列"))
        self.pushButton_5.setText(_translate("MainWindow", "显示所有行"))

    def openfile(self):
        self.statusbar.showMessage('')
        try:
            openfile_name = QFileDialog.getOpenFileName(None, '选择文件', 'E:\\科技项目相关\\201903', 'Excel files(*.xlsx , *.xls)')
            self.lineEdit.setText(openfile_name[0])
        except:
            self.statusbar.showMessage('选择文件失败')

    def creat_table_show(self, index):
        self.statusbar.showMessage('')
        self.label_6.setText("0.0")
        self.last_check_state = []
        try:
            data = xlrd.open_workbook(self.lineEdit.text())
            table = data.sheet_by_index(index)
            cols = table.ncols
            rows = table.nrows

            self.tableWidget.setRowCount(rows)
            self.tableWidget.setColumnCount(cols+1)
            self.table_show_colnums()

            if index == 0:
                self.comboBox.clear()
                names = data.sheet_names()
                for name in names:
                    self.comboBox.addItem(name)
            self.comboBox2.clear()
            self.comboBox3.clear()
            for i in range(1, cols+2):
                self.comboBox2.addItem(str(i))
                self.comboBox3.addItem(str(i))
            for i in range(rows):
                if i == 0:
                    checkbox = QTableWidgetItem('All')
                else:
                    checkbox = QTableWidgetItem(str(i + 1))
                checkbox.setFlags(QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
                checkbox.setCheckState(QtCore.Qt.Unchecked)
                self.tableWidget.setItem(i, 0, checkbox)
                self.tableWidget.setColumnWidth(0, 50)
                self.last_check_state.append(QtCore.Qt.Unchecked)
                for j in range(cols):
                    value = table.cell(i, j).value
                    self.tableWidget.setItem(i, j+1, QTableWidgetItem(str(value)))
            self.tableWidget.editTriggers()
        except:
            self.statusbar.showMessage('打开文件异常')

    def table_hidden_changed(self, index):
        if index == 0 or index > self.tableWidget.columnCount():
            self.statusbar.showMessage('该列无法被隐藏')
        else:
            self.statusbar.showMessage('')
            if self.tableWidget.isColumnHidden(index):
                self.tableWidget.showColumn(index)
            else:
                self.tableWidget.hideColumn(index)

    def table_vertical_state_changed(self, index):
        if index == 0 or index > self.tableWidget.rowCount():
            self.statusbar.showMessage('该行无法被隐藏')
        else:
            self.statusbar.showMessage('')
            if self.tableWidget.isRowHidden(index):
                self.tableWidget.showRow(index)
            else:
                self.tableWidget.hideRow(index)

    def table_show_colnums(self):
        for i in range(self.tableWidget.columnCount()):
            self.tableWidget.showColumn(i)

    def table_show_verticals(self):
        for i in range(self.tableWidget.rowCount()):
            self.tableWidget.showRow(i)

    def get_sum(self, item):
        self.statusbar.showMessage('')
        if item.column() != 0:
            self.statusbar.showMessage('该行无值无法计算')
            return
        try:
            sum_count = self.label_6.text()
            sum_count = float(sum_count)
            sum_count = round(sum_count, 2)
        except:
            self.statusbar.showMessage('总金额错误')
            return
        if item.checkState() == self.last_check_state[item.row()]:
            return
        else:
            if item.row() == 0:
                if item.checkState() == QtCore.Qt.Checked:
                    self.last_check_state[0] = QtCore.Qt.Checked
                    for i in range(1, self.tableWidget.rowCount()):
                        try:
                            if self.last_check_state[i] == QtCore.Qt.Unchecked:
                                self.last_check_state[i] = QtCore.Qt.Checked
                                self.tableWidget.item(i, 0).setCheckState(QtCore.Qt.Checked)
                                cell_item = self.tableWidget.item(i, self.comboBox3.currentIndex())
                                value = float(cell_item.text())
                                value = round(value, 2)
                                sum_count += value
                        except:
                            continue
                else:
                    self.last_check_state[0] = QtCore.Qt.Unchecked
                    for i in range(1, self.tableWidget.rowCount()):
                        try:
                            if self.last_check_state[i] == QtCore.Qt.Checked:
                                self.last_check_state[i] = QtCore.Qt.Unchecked
                                self.tableWidget.item(i, 0).setCheckState(QtCore.Qt.Unchecked)
                                cell_item = self.tableWidget.item(i, self.comboBox3.currentIndex())
                                value = float(cell_item.text())
                                value = round(value, 2)
                                sum_count -= value
                                sum_count = round(sum_count, 2)
                        except:
                            continue
            else:
                cell_item = self.tableWidget.item(item.row(), self.comboBox3.currentIndex())
                try:
                    value = float(cell_item.text())
                    value = round(value, 2)
                    if item.checkState() == QtCore.Qt.Checked:
                        self.last_check_state[item.row()] = QtCore.Qt.Checked
                        sum_count += value
                    else:
                        self.last_check_state[item.row()] = QtCore.Qt.Unchecked
                        sum_count -= value
                except:
                    self.statusbar.showMessage('值无法计算')
            self.label_6.setText(str(sum_count))

    def save_result(self):
        try:
            with open(".\\result.txt", 'w') as fd:
                for i in range(self.tableWidget.rowCount()):
                    if self.last_check_state[i] == QtCore.Qt.Checked:
                        line_content = "'"
                        for j in range(1, self.tableWidget.columnCount()):
                            if not self.tableWidget.isColumnHidden(j):
                                print_item = self.tableWidget.item(i, j)
                                line_content += print_item.text()
                                line_content += '&'
                        line_content += '\n'
                        fd.write(line_content)
            self.statusbar.showMessage('保存文件到result.txt成功')
        except:
            self.statusbar.showMessage('保存文件错误')

