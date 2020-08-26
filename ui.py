# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'mech-01.ui'
#
# Created by: PyQt5 UI code generator 5.13.2
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1127, 560)
        font = QtGui.QFont()
        font.setPointSize(16)
        MainWindow.setFont(font)
        MainWindow.setStyleSheet("background-color: rgb(106, 106, 106);")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(270, 10, 841, 21))
        font = QtGui.QFont()
        font.setFamily("Ubuntu")
        font.setPointSize(14)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.label.setFont(font)
        self.label.setStyleSheet("font: 14pt \"Ubuntu\";\n"
"color: rgb(230, 230, 230);\n"
"color: rgb(255, 255, 255);")
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setObjectName("label")
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setGeometry(QtCore.QRect(10, 10, 241, 451))
        self.frame.setStyleSheet("background-color: rgb(117, 117, 117);\n"
"border-radius:10px")
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.pushButton = QtWidgets.QPushButton(self.frame)
        self.pushButton.setGeometry(QtCore.QRect(10, 10, 221, 41))
        self.pushButton.setStyleSheet("font: 10pt \"Ubuntu\";\n"
"background-color: rgb(84, 84, 84);\n"
"color: rgb(255, 255, 255);\n"
"border-style: outset;\n"
"border-radius: 10px")
        self.pushButton.setObjectName("pushButton")
        self.listWidget = QtWidgets.QListWidget(self.frame)
        self.listWidget.setGeometry(QtCore.QRect(10, 60, 221, 381))
        self.listWidget.setObjectName("listWidget")
        self.tableView = QtWidgets.QTableView(self.centralwidget)
        self.tableView.setGeometry(QtCore.QRect(280, 60, 821, 461))
        font = QtGui.QFont()
        font.setFamily("Ubuntu")
        font.setPointSize(8)
        self.tableView.setFont(font)
        self.tableView.setStyleSheet("background-color: rgb(103, 103, 103);\n"
"background-color: rgb(98, 98, 98);\n"
"color: rgb(243, 243, 243);\n"
"border-radius:10px\n"
"\n"
"")
        self.tableView.setObjectName("tableView")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(10, 480, 241, 41))
        self.pushButton_2.setStyleSheet("font: 10pt \"Ubuntu\";\n"
"background-color: rgb(84, 84, 84);\n"
"color: rgb(255, 255, 255);\n"
"border-style: outset;\n"
"border-radius: 10px")
        self.pushButton_2.setObjectName("pushButton_2")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Расчет опор на механическую прочность"))
        self.label.setText(_translate("MainWindow", "Расчет опор на механическую прочность"))
        self.pushButton.setText(_translate("MainWindow", "выбрать папку с файлами"))
        self.pushButton_2.setText(_translate("MainWindow", "выполнить расчёт"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
