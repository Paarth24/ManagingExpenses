from openpyxl import Workbook, load_workbook
from PyQt5.QtCore import QSize
from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QPushButton, QVBoxLayout, QWidget

import pathlib
import os

from BuildDirPath import BuildDirPath
#from BuildDirWindow import BuildDirWindow 
from ExcelDisplayList import ExcelDisplayList
from Functions import GetExcelFileName
from MergeWindow import MergeWindow
from ExcelDataBase import ExcelDataBase  

class MainWindow(QMainWindow, ExcelDataBase, ExcelDisplayList):
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
        mainLayout = QVBoxLayout()
#-------------------Initialization--------------------------- 
#-------------------Formatting---------------------------         
        mainLayout.addWidget(mergeButton)
        mainLayout.addWidget(addButton)
        mainLayout.addWidget(buildButton)
        mainLayout.addWidget(BuildDirPath.buildDir)
        mainLayout.addWidget(ExcelDisplayList.excelDisplay)

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
        excelNames = QFileDialog.getOpenFileNames(self, "Add Excels", "C:/", "Excel Files (*.xlsx)")
        if (excelNames != None):
            for i in range(0, len(excelNames) - 1):
                ExcelDataBase.excelList.append(excelNames[i])
            print(ExcelDataBase.excelList)
#-------------------Display---------------------------
            ExcelDisplayList.excelDisplay.clear()      
            for i in range(0, len(ExcelDataBase.excelList)):
                userExcelFile = GetExcelFileName(ExcelDataBase.excelList[i])
                print(userExcelFile)
                ExcelDisplayList.excelDisplay.addItem(userExcelFile)
#-------------------Display---------------------------        
        
    
    def ChangingBuildDirectory(self):
        
        userPath = QFileDialog.getExistingDirectory(self, "Build Location", "C:/")
        userDefinedPath = userPath
        rawBuildPath = pathlib.PureWindowsPath(userDefinedPath)
        pureBuildPath = str(rawBuildPath.as_posix())
        os.chdir(pureBuildPath)
        BuildDirPath.buildDir.setText(os.getcwd())
        
    def MergingAddedFiles(self):
        self.mergeWindow = MergeWindow()
        self.mergeWindow.show()
    
    def closeEvent(self, event):
        QApplication.closeAllWindows()
#-------------------Functions---------------------------



