from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMainWindow, QApplication, QDialog, QWidget
from MoveData import  moveData, TheSheets
from MoveData_97 import moveData97, TheSheet97
import os, time
from PyQt5.QtCore import Qt
from pathlib import PurePath as p
from streamlit import caching
caching.clear_cache()

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(279, 145)
        MainWindow.setAcceptDrops(False)
        MainWindow.setStyleSheet("background-color: rgb(255, 255, 255);")
        MainWindow.setWindowFilePath("")
        MainWindow.setToolButtonStyle(QtCore.Qt.ToolButtonIconOnly)
        MainWindow.setTabShape(QtWidgets.QTabWidget.Rounded)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.widget = QtWidgets.QWidget(self.centralwidget)
        self.widget.setObjectName("widget")
        self.gridLayout = QtWidgets.QGridLayout(self.widget)
        self.gridLayout.setObjectName("gridLayout")
        self.pushButton_loadCopy = QtWidgets.QPushButton(self.widget)
        self.pushButton_loadCopy.setStyleSheet(
"    background:linear-gradient(to bottom, #ffec64 5%, #ffab23 100%);\n"
"    background-color:#ffec64;\n"
"    border-radius:9px;\n"
"    border:1px solid #ffaa22;\n"
"    color:#333333;\n"
"    font-family:Arial;\n"
"    font-size:14px;\n"
"    font-weight:bold;\n"
"    padding:5px 5px;\n"
"    text-decoration:none;\n")
        self.pushButton_loadCopy.setObjectName("pushButton_loadCopy")
        self.gridLayout.addWidget(self.pushButton_loadCopy, 0, 0, 1, 1)
        self.label_Copy = QtWidgets.QLabel(self.widget)
        self.label_Copy.setStyleSheet("")
        self.label_Copy.setText("")
        self.label_Copy.setObjectName("label_Copy")
        self.gridLayout.addWidget(self.label_Copy, 0, 1, 1, 1)
        self.pushButton_Stop = QtWidgets.QPushButton(self.widget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(1)
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        font.setStrikeOut(False)
        self.pushButton_Stop.setFont(font)
        self.pushButton_Stop.setStatusTip("")
        self.pushButton_Stop.setStyleSheet(
"    background:linear-gradient(to bottom, #ffec64 5%, #ffab23 100%);\n"
"    background-color:#ffec64;\n"
"    border-radius:9px;\n"
"    border:1px solid #ffaa22;\n"
"    color:#333333;\n"
"    font-family:Arial;\n"
"    font-size:14px;\n"
"    font-weight:bold;\n"
"    padding:5px 5px;\n"
"    text-decoration:none;\n")
        self.pushButton_Stop.setObjectName("pushButton_Stop")
        self.gridLayout.addWidget(self.pushButton_Stop, 1, 2, 1, 1)
        self.pushButton_Start = QtWidgets.QPushButton(self.widget)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(1)
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        font.setStrikeOut(False)
        self.pushButton_Start.setFont(font)
        self.pushButton_Start.setStatusTip("")
        self.pushButton_Start.setStyleSheet(
"    background:linear-gradient(to bottom, #ffec64 5%, #ffab23 100%);\n"
"    background-color:#ffec64;\n"
"    border-radius:9px;\n"
"    border:1px solid #ffaa22;\n"
"    color:#333333;\n"
"    font-family:Arial;\n"
"    font-size:14px;\n"
"    font-weight:bold;\n"
"    padding:5px 5px;\n"
"    text-decoration:none;\n")
        self.pushButton_Start.setObjectName("pushButton_Start")
        self.gridLayout.addWidget(self.pushButton_Start, 0, 2, 1, 1)
        self.pushButton_laodPaste = QtWidgets.QPushButton(self.widget)
        self.pushButton_laodPaste.setStyleSheet(
"    background:linear-gradient(to bottom, #ffec64 5%, #ffab23 100%);\n"
"    background-color:#ffec64;\n"
"    border-radius:9px;\n"
"    border:1px solid #ffaa22;\n"
"    color:#333333;\n"
"    font-family:Arial;\n"
"    font-size:14px;\n"
"    font-weight:bold;\n"
"    padding:5px 5px;\n"
"    text-decoration:none;\n")
        self.pushButton_laodPaste.setObjectName("pushButton_laodPaste")
        self.gridLayout.addWidget(self.pushButton_laodPaste, 1, 0, 1, 1)
        self.progressBar = QtWidgets.QProgressBar(self.widget)
        self.progressBar.setStyleSheet("")
        self.progressBar.setProperty("value", 0)
        self.progressBar.setTextVisible(False)
        self.progressBar.setObjectName("progressBar")
        self.gridLayout.addWidget(self.progressBar, 2, 0, 1, 3)
        self.label_Paste = QtWidgets.QLabel(self.widget)
        self.label_Paste.setStyleSheet("")
        self.label_Paste.setText("")
        self.label_Paste.setObjectName("label_Paste")
        self.gridLayout.addWidget(self.label_Paste, 1, 1, 1, 1)
        self.gridLayout_2.addWidget(self.widget, 0, 0, 1, 1)
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
        self.pushButton_Stop.setToolTip(_translate("MainWindow", "press to exit"))
        self.pushButton_Stop.setText(_translate("MainWindow", "Exit"))
        self.pushButton_Start.setToolTip(_translate("MainWindow", "press to copy"))
        self.pushButton_Start.setText(_translate("MainWindow", "Copy"))
        self.pushButton_laodPaste.setToolTip(_translate("MainWindow", "press to select a paste file"))
        self.pushButton_laodPaste.setText(_translate("MainWindow", "Paste File"))
        self.statusbar.setToolTip(_translate("MainWindow", "here"))

    def loading_(self):
        self.statusbar.showMessage("Loading.....")

    def exiter(self):
        sys.exit(1)

    def myFile(self):
        self.gg , _ = QtWidgets.QFileDialog.getOpenFileName(MainWindow)
        self.label_Paste.setText(os.path.basename(self.gg))
        return self.gg

    def myFolder(self):
        try:
            self.hh= QFileDialog.getExistingDirectory(MainWindow)
            self.label_Copy.setText(os.path.basename(self.hh))
            os.chdir(self.hh)
            ff= os.listdir(self.hh)
            f=[sf for sf in ff if sf.endswith('.xls') or sf.endswith('.xlsx')]
            no_of_Sheets=0
            no_of_Files=0
            for copyFileName in f:
                if copyFileName.endswith('.xlsx'):
                    _ , wss= TheSheets(copyFileName)
                    no_of_Files= no_of_Files+1
                    for _ in wss:
                        no_of_Sheets= no_of_Sheets+1
                if copyFileName.endswith('.xls'):
                    _ , wss= TheSheet97(copyFileName)
                    no_of_Files= no_of_Files+1
                    for _ in wss:
                        no_of_Sheets= no_of_Sheets+1
            self.progressBar.setMaximum(no_of_Sheets)
            self.statusbar.showMessage("{} files are found...".format( no_of_Files))
            return self.hh
        except:
            self.label_Copy.setText("Please Enter correct folder")
        
    
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
        self.statusbar.showMessage("In progress....")
        os.chdir(copyFolder)
        ff= os.listdir(copyFolder)
        f97=[sf for sf in ff if sf.endswith('.xls')]
        f97.sort(key= os.path.getmtime)
        f= [sf for sf in ff if sf.endswith('.xlsx')]
        f.sort(key= os.path.getmtime)
        ind2=0
        message=""
        message1=""
        successList=[]
        failList=[]

        # Copy sheets of .xls format into the paste file .xlsx
        for copyFileName in f97:
            _ , wss= TheSheet97(copyFileName)
            for copyFileSheet in wss:
                ind2= ind2+1
                self.increaseIt(ind2)
                messageInProgress= "In progress.... Copy {} ({})".format(copyFileName, copyFileSheet)
                self.statusbar.showMessage(messageInProgress)
                result= moveData97(copyFileName, copyFileSheet, pasteFileName)
                print(result)
                if result ==True:
                    successList.append([copyFileName,copyFileSheet])
                else:
                    failList.append([copyFileName,copyFileSheet])


        # Copy sheets of .xlsx format into the paste file .xlsx
        for copyFileName in f:
            _ , wss= TheSheets(copyFileName)
            for copyFileSheet in wss:
                ind2= ind2+1
                self.increaseIt(ind2)
                messageInProgress= "In progress.... Copy {} ({})".format(copyFileName, copyFileSheet)
                self.statusbar.showMessage(messageInProgress)
                try:    
                    time.sleep(0.05)

                    result=moveData(copyFileName, copyFileSheet, pasteFileName)
                    
                    if result== True:
                        successList.append([copyFileName,copyFileSheet])
                    else:
                        failList.append([copyFileName,copyFileSheet])
                except:
                    failList.append([copyFileName,copyFileSheet])
        
        self.increaseIt(0)
        finish_time= time.perf_counter()
        if successList:
            message= 'Copied {} Sheets in {} Seconds'.format( len(successList), round(finish_time- start_time, 2))
        if failList:
            message1= "\n"+"Failed to copy {} Sheets".format(failList)    
        self.statusbar.showMessage(message +message1)
        mList= os.path.basename(copyFolder) + " : " + message+ message1+ "\n"
        self.exportToLastStatus(mList)

    def exportToLastStatus(self, mList):
        pasteDirectory= os.path.dirname(self.gg)
        LogFileLocation= p(pasteDirectory). joinpath('log.txt')
        if mList:
            writer= open(LogFileLocation,'a')
            writer.write(mList)
            writer.close()


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())