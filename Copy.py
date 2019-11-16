# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'FileName1.ui'
#
# Created by: PyQt5 UI code generator 5.13.0
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMainWindow, QApplication
from MoveData import  moveData, TheSheets
import os, time

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(258, 138)
        MainWindow.setAcceptDrops(False)
        MainWindow.setStyleSheet("")
        MainWindow.setWindowFilePath("")
        MainWindow.setToolButtonStyle(QtCore.Qt.ToolButtonIconOnly)
        MainWindow.setTabShape(QtWidgets.QTabWidget.Rounded)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.pushButton_loadCopy = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_loadCopy.setStyleSheet("")
        self.pushButton_loadCopy.setObjectName("pushButton_loadCopy")
        self.gridLayout.addWidget(self.pushButton_loadCopy, 0, 0, 1, 1)
        self.label_Copy = QtWidgets.QLabel(self.centralwidget)
        self.label_Copy.setStyleSheet("")
        self.label_Copy.setText("")
        self.label_Copy.setObjectName("label_Copy")
        self.gridLayout.addWidget(self.label_Copy, 0, 1, 1, 1)
        self.pushButton_Start = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_Start.setFont(font)
        self.pushButton_Start.setStatusTip("")
        self.pushButton_Start.setStyleSheet("")
        self.pushButton_Start.setObjectName("pushButton_Start")
        self.gridLayout.addWidget(self.pushButton_Start, 0, 2, 1, 1)
        self.pushButton_laodPaste = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_laodPaste.setObjectName("pushButton_laodPaste")
        self.gridLayout.addWidget(self.pushButton_laodPaste, 1, 0, 1, 1)
        self.label_Paste = QtWidgets.QLabel(self.centralwidget)
        self.label_Paste.setStyleSheet("")
        self.label_Paste.setText("")
        self.label_Paste.setObjectName("label_Paste")
        self.gridLayout.addWidget(self.label_Paste, 1, 1, 1, 1)
        self.pushButton_Stop = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_Stop.setFont(font)
        self.pushButton_Stop.setStatusTip("")
        self.pushButton_Stop.setObjectName("pushButton_Stop")
        self.gridLayout.addWidget(self.pushButton_Stop, 1, 2, 1, 1)
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setTextVisible(False)
        self.progressBar.setObjectName("progressBar")
        self.gridLayout.addWidget(self.progressBar, 2, 0, 1, 3)
        self.label_Status = QtWidgets.QLabel(self.centralwidget)
        self.label_Status.setStyleSheet("")
        self.label_Status.setText("")
        self.label_Status.setObjectName("label_Status")
        self.gridLayout.addWidget(self.label_Status, 3, 0, 1, 3)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.pushButton_laodPaste.clicked.connect( self.myFile)
        self.pushButton_loadCopy.clicked.connect( self.myFolder)
        self.pushButton_Start.clicked.connect( self.Copy_btn)
        self.pushButton_Stop.clicked.connect(self.exiter)
        self.pushButton_loadCopy.pressed.connect(self.loading_)
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Activity Copier"))
        self.pushButton_loadCopy.setToolTip(_translate("MainWindow", "press to load activity folder"))
        self.pushButton_loadCopy.setText(_translate("MainWindow", "Copy Folder"))
        self.pushButton_Start.setToolTip(_translate("MainWindow", "press to copy"))
        self.pushButton_Start.setText(_translate("MainWindow", "Copy"))
        self.pushButton_laodPaste.setToolTip(_translate("MainWindow", "press to select a paste file"))
        self.pushButton_laodPaste.setText(_translate("MainWindow", "Paste File"))
        self.pushButton_Stop.setToolTip(_translate("MainWindow", "press to exit"))
        self.pushButton_Stop.setText(_translate("MainWindow", "Exit"))

    def loading_(self):
        self.label_Status.setText("Loading.....")

    def exiter(self):
        sys.exit(1)

    def myFile(self):
        self.gg , _ = QtWidgets.QFileDialog.getOpenFileName(MainWindow)
        self.label_Paste.setText(os.path.basename(self.gg))
        return self.gg

    def myFolder(self):
        self.hh= QFileDialog.getExistingDirectory(MainWindow)
        self.label_Copy.setText(os.path.basename(self.hh))
        os.chdir(self.hh)
        ff= os.listdir(self.hh)
        f=[sf for sf in ff if sf.endswith('.xls') or sf.endswith('.xlsx')]
        no_of_Sheets=0
        no_of_Files=0
        for copyFileName in f:
            _ , wss= TheSheets(copyFileName)
            no_of_Files= no_of_Files+1
            for _ in wss:
                no_of_Sheets= no_of_Sheets+1
        self.progressBar.setMaximum(no_of_Sheets)
        self.label_Status.setText("{} files are found...".format( no_of_Files))
        return self.hh
    
    def Copy_btn(self):
        try:
            self.hh
            try:
                self.gg
                try:
                    self.start(self.gg, self.hh)
                except Exception:
                    pass
            except:
                self.label_Paste.setText("Please enter the File name")
        except:
            self.label_Copy.setText("Please enter the folder name")

    def increaseIt(self, value):
        self.progressBar.setValue(value)

    def start(self, pasteFileName, copyFolder):
        start_time= time.perf_counter()
        self.label_Status.setText("In progress....")
        os.chdir(copyFolder)
        ff= os.listdir(copyFolder)
        f=[sf for sf in ff if sf.endswith('.xls') or sf.endswith('xlsx')]
        ind2=0
        message=""
        message1=""
        successList=[]
        failList=[]
        for copyFileName in f:
            _ , wss= TheSheets(copyFileName)
            for copyFileSheet in wss:
                try:
                    messageInProgress= "In progress.... Copy {} ({})".format(copyFileName, copyFileSheet)
                    self.label_Status.setText(messageInProgress)
                    moveData(copyFileName, copyFileSheet, pasteFileName)
                    ind2= ind2+1
                    self.increaseIt(ind2)
                    if copyFileName not in successList:
                        successList.append(copyFileName)
                except:
                    if copyFileName not in failList:
                        failList.append(copyFileName)
        self.increaseIt(0)
        finish_time= time.perf_counter()
        if failList:
            message1= "\n"+"Failed to copy {} files".format(len(failList))
        if successList:
            message= 'Copied {} files in {} Second(s)'.format( len(successList), round(finish_time- start_time, 2))
        self.label_Status.setText(message +message1)


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
