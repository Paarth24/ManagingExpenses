from openpyxl import Workbook, load_workbook
from PyQt5.QtCore import QSize
from PyQt5.QtWidgets import QApplication, QBoxLayout, QFileDialog, QGridLayout, QHBoxLayout, QLabel, QLineEdit, QListView, QListWidget, QMainWindow, QPushButton, QVBoxLayout, QWidget
import os
import pathlib
from Functions import GetExcelFileName
from MergeWindow import MergeWindow
from ExcelDataBase import ExcelDataBase

class MainWindow(QMainWindow, ExcelDataBase):
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
            ExcelDataBase.excelList.append(excelName[0])
            
#-------------------Display---------------------------
            self.displayExcel.clear()      
            for i in range(0, len(ExcelDataBase.excelList)):
                userExcelFile = GetExcelFileName(ExcelDataBase.excelList[i])
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
    ExcelDataBase().initialize()
    app = QApplication([])
    window = MainWindow()
    window.show()

app.exec_()
