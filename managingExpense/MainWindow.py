from PyQt5.QtCore import QSize
from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QPushButton, QVBoxLayout, QWidget

import pathlib
import os

from Functions import GetExcelFileName
from MergeWindow import MergeWindow
from Resources import RESOURCES  

class MainWindow(QMainWindow, RESOURCES):
#-------------------MAINWINDOW---------------------------
    def __init__(self):
        super().__init__()

#-------------------configuration---------------------------
        self.setWindowTitle("Managing Expenses")
        self.setMinimumSize(QSize(500, 300))
#-------------------configuration---------------------------
#-------------------Initialization---------------------------        
        mainLayout = QVBoxLayout()
        mergeButton = QPushButton("Merge")
        addButton = QPushButton("Add")
        buildButton = QPushButton("Build Folder")       
#-------------------Initialization--------------------------- 
#-------------------Formatting---------------------------         
        mainLayout.addWidget(mergeButton)
        mainLayout.addWidget(addButton)
        mainLayout.addWidget(buildButton)
        mainLayout.addWidget(RESOURCES.buildDirLabel)
        mainLayout.addWidget(RESOURCES.excelDisplay)

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
            for i in range(0, len(excelNames[0])):
                RESOURCES.excelList.append(excelNames[0][i])
#-------------------Display---------------------------
            RESOURCES.excelDisplay.clear()      
            for i in range(0, len(RESOURCES.excelList)):
                userExcelFile = GetExcelFileName(RESOURCES.excelList[i])
                RESOURCES.excelDisplay.addItem(userExcelFile)
#-------------------Display---------------------------        
        
    
    def ChangingBuildDirectory(self):
        
        userPath = QFileDialog.getExistingDirectory(self, "Build Location", "C:/")
        userDefinedPath = userPath
        rawBuildPath = pathlib.PureWindowsPath(userDefinedPath)
        pureBuildPath = rawBuildPath.as_posix()
        os.chdir(pureBuildPath)
        RESOURCES.buildDirLabel.setText("Build Folder - {}".format(os.getcwd()))
        
    def MergingAddedFiles(self):
        self.mergeWindow = MergeWindow()
        self.mergeWindow.show()
    
    def closeEvent(self, event):
        QApplication.closeAllWindows()
#-------------------Functions---------------------------



