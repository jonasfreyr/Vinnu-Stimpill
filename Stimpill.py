# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'D:\Python\vinnu-stimpill\gluggi.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

import datetime, openpyxl



YEAR = datetime.date.today().year

SHEET = datetime.datetime.now()
SHEET = str(datetime.datetime.strftime(SHEET, "%b"))

try:
    wb = openpyxl.load_workbook(str(YEAR)+".xlsx")

    try:
        e = wb[SHEET]
    except:
        e = wb.create_sheet(SHEET)

        e["A1"] = "Dags"
        e["B1"] = "Inn"
        e["C1"] = "Út"
        e["D1"] = "Tímar"

except:
    wb = openpyxl.Workbook()

    e = wb.active
    e.title = SHEET
    e["A1"] = "Dags"
    e["B1"] = "Inn"
    e["C1"] = "Út"
    e["D1"] = "Tímar"

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 600)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(50, 210, 121, 31))
        self.pushButton.setObjectName("pushButton")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(270, 90, 241, 71))
        font = QtGui.QFont()
        font.setPointSize(29)
        self.label.setFont(font)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setObjectName("label")
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(630, 210, 121, 31))
        self.pushButton_3.setObjectName("pushButton_3")
        self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget.setEnabled(True)
        self.tableWidget.setGeometry(QtCore.QRect(50, 260, 551, 291))
        self.tableWidget.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.listWidget = QtWidgets.QListWidget(self.centralwidget)
        self.listWidget.setGeometry(QtCore.QRect(630, 260, 121, 291))
        self.listWidget.setObjectName("listWidget")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(350, 210, 91, 31))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.label_2.setFont(font)
        self.label_2.setText("")
        self.label_2.setObjectName("label_2")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        MainWindow.setFixedSize(MainWindow.width(), MainWindow.height())

        MainWindow.setWindowIcon(QtGui.QIcon('icon.ico'))

        self.pushButton.clicked.connect(self.inn)
        self.pushButton_3.clicked.connect(self.out)

        self.inni = False

        self.columns = {"A": 0, "B": 1, "C": 2, "D": 3}

        self.time = None

        self.row = e.max_row+1

        self.tableWidget.setColumnCount(4)
        self.tableWidget.setRowCount(32)

        self.label_2.setText(e.title)

        self.load_data(e, True)

        self.selected = None

        self.listWidget.itemClicked.connect(self.change_sheet)

        for name in wb.sheetnames:
            self.listWidget.addItem(name)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Stimpill"))
        self.pushButton.setText(_translate("MainWindow", "Inn"))
        self.label.setText(_translate("MainWindow", "Inn eða út?"))
        self.pushButton_3.setText(_translate("MainWindow", "Út"))

    def change_sheet(self, s):
        self.tableWidget.clear()

        try:
            self.label_2.setText(s.text())
            self.load_data(wb[s.text()])

        except:
            self.label_2.setText(s)
            self.load_data(wb[s])

    def save(self):
        wb.save(str(YEAR)+".xlsx")

    def load_data(self, e, inital=False):
        inn_count = 0
        out_count = 0
        for c in "ABCD":
            column = e[c]

            for a in range(len(column)):
                if column[a].value != None:
                    self.tableWidget.setItem(a, self.columns[c], QtWidgets.QTableWidgetItem(str(column[a].value)))

                    if c == "B" and column[a].value != "Inn":
                        time = str(column[a].value)

                    elif c == "A" and column[a].value != "Dags":
                        date = str(column[a].value)

                    if c == "B":
                        inn_count += 1

                    elif c == "C":
                        out_count += 1

        if inital:
            if inn_count > out_count:
                self.inni = True
                self.row -= 1

                self.time = date + " " + time

                self.time = datetime.datetime.strptime(self.time, "%d.%m.%Y %H:%M:%S")

    def inn(self):
        if not self.inni:
            d = datetime.datetime.now()

            time = d.strftime("%H:%M:%S")

            self.time = d

            self.inni = True

            date = d.strftime("%d.%m.%Y")

            e.cell(row=self.row, column=2).value = time

            e.cell(row=self.row, column=1).value = date

            self.save()

            self.change_sheet(e.title)

    def out(self):
        if self.inni:
            d = datetime.datetime.now()

            time = d.strftime("%H:%M:%S")

            hours = d - self.time
            secs = hours.total_seconds()

            hours = round(secs / 3600, 2)

            self.inni = False

            e.cell(row=self.row, column=3).value = time

            e.cell(row=self.row, column=4).value = hours

            self.row += 1

            self.save()

            self.change_sheet(e.title)

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

