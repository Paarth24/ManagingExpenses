from PyQt5.QtGui import QCloseEvent
from openpyxl import Workbook, load_workbook
from PyQt5.QtCore import QSize
from PyQt5.QtWidgets import QApplication, QBoxLayout, QFileDialog, QGridLayout, QHBoxLayout, QLabel, QLineEdit, QListView, QListWidget, QMainWindow, QPushButton, QVBoxLayout, QWidget
import os
import pathlib
from BuildDirWindow import BuildDirWindow
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
        self.buildDirectory = QLabel(os.getcwd())       
        mainLayout = QVBoxLayout()
#-------------------Initialization--------------------------- 
#-------------------Formatting---------------------------         
        mainLayout.addWidget(mergeButton)
        mainLayout.addWidget(addButton)
        mainLayout.addWidget(buildButton)
        mainLayout.addWidget(self.buildDirectory)
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
        excelName = QFileDialog.getOpenFileName(self, "Add Excels", "C:/", "Excel Files (*.xlsx)")
        if (excelName != ("","")):
            #Number of excels to be merged
            ExcelDataBase.excelList.append(excelName[0])         
#-------------------Display---------------------------
            self.displayExcel.clear()      
            for i in range(0, len(ExcelDataBase.excelList)):
                userExcelFile = GetExcelFileName(ExcelDataBase.excelList[i])
                ExcelDisplayList.excelDisplay.addItem(userExcelFile)
#-------------------Display---------------------------        
        
    
    def ChangingBuildDirectory(self):
        #Changing build location
        self.buildDirWindow = BuildDirWindow()
        self.buildDirWindow.show()
        
    def MergingAddedFiles(self):
        self.mergeWindow = MergeWindow()
        self.mergeWindow.show()
    
    def closeEvent(self, event):
        if(self.mergeWindow.isVisible()):
            self.mergeWindow.close()
#-------------------Functions---------------------------



