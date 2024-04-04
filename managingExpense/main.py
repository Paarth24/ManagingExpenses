from openpyxl import Workbook, load_workbook
from PyQt5.QtCore import QSize
from PyQt5.QtWidgets import QApplication, QBoxLayout, QFileDialog, QGridLayout, QHBoxLayout, QLabel, QLineEdit, QListView, QListWidget, QMainWindow, QPushButton, QVBoxLayout, QWidget
import os
import pathlib
from MergeWindow import MergeWindow
from GlobalVar import GlobalVar

def RemovingFileExt(fileName):
    fileNameNoExt = ""

    for i in range(0, len(fileName)):
        if(fileName[i] == "."):
            break
        else:
            fileNameNoExt = fileNameNoExt + fileName[i]

    return(fileNameNoExt)

def GetExcelFileName(path):
    count = 0
    for i in range(0, len(path)):
        if(path[i] == "/"):
           count = count + 1

    file = ""
    for i in range(0, len(path)):
        if(path[i] == "/"):
           count = count - 1


        if(count == -1):
            file = file + path[i]

        if(count == 0):
            count = -1
            
    return(file)


class MainWindow(QMainWindow):
#-------------------MAINWINDOW---------------------------
    def __init__(self):
        super().__init__()
        
#-------------------configuration---------------------------
        self.setWindowTitle("Test App")
        self.setMinimumSize(QSize(500, 300))
#-------------------configuration---------------------------
#-------------------Initialization---------------------------        
        mergeButton = QPushButton("Merge")
        addButton = QPushButton("Add")
        buildButton = QPushButton("Build Folder")
        self.buildDirectory = QLabel(os.getcwd())
        self.displayExcel = QListWidget()

        self.excelList = []
        
        mainLayout = QVBoxLayout()
#-------------------Initialization--------------------------- 
#-------------------Formatting---------------------------         
        mainLayout.addWidget(mergeButton)
        mainLayout.addWidget(addButton)
        mainLayout.addWidget(buildButton)
        mainLayout.addWidget(self.buildDirectory)
        mainLayout.addWidget(self.displayExcel)

        mainContainer = QWidget()
        mainContainer.setLayout(mainLayout)
        
        self.setMenuWidget(mainContainer)
#-------------------Formatting---------------------------
#-------------------Signals---------------------------
        addButton.clicked.connect(self.AddingExcelFiles)
        buildButton.clicked.connect(self.ChangingBuildDirectory)
        mergeButton.clicked.connect(self.MergingAddedFiles)
#-------------------Signals---------------------------
#-------------------MAINWINDOW--------------------------- 
       
#-------------------Functions--------------------------- 
    def AddingExcelFiles(self):
        excelName = QFileDialog.getOpenFileName(self, "Add Excels", "C:/", "Excel Files (*.xlsx)")
        if (excelName != ("","")):
            #Number of excels to be merged
            GlobalVar().ExcelList().append(excelName[0])
            
#-------------------Display---------------------------
            self.displayExcel.clear()      
            for i in range(0, len(GlobalVar().ExcelList())):
                userExcelFile = GetExcelFileName(GlobalVar().ExcelList()[i])
                self.displayExcel.addItem(userExcelFile)
#-------------------Display---------------------------        
        
    
    def ChangingBuildDirectory(self):
        #Changing build location
        userPath = QFileDialog.getExistingDirectory(self, "Build Location", "C:/")
        userDefinedPath = userPath
        rawBuildPath = pathlib.PureWindowsPath(userDefinedPath)
        pureBuildPath = str(rawBuildPath.as_posix())
        os.chdir(pureBuildPath)
        self.buildDirectory.setText(os.getcwd())
        
    def MergingAddedFiles(self):
        self.mergeWindow = MergeWindow()
        self.mergeWindow.show()
        
#-------------------Functions---------------------------          


if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()

app.exec_()
