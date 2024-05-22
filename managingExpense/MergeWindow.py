import datetime
import os
import pathlib
from openpyxl import Workbook, load_workbook
from PyQt5.QtWidgets import QLabel, QLineEdit, QMainWindow, QPushButton, QVBoxLayout, QWidget

from ErrorWindow import ErrorWindow
from Resources import RESOURCES
from Functions import GetExcelFileName, AddingFileExt, IfValue, DATETIME
from AfterMerge import AfterMerge




class MergeWindow(ErrorWindow, RESOURCES):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Merge Window")
        self.setMinimumSize(250, 100)
        
        RESOURCES.buildFileNameNoExt = "Final_Bank_Statement"
        RESOURCES.buildFileNameExt = AddingFileExt(self.buildFileNameNoExt)

        layout = QVBoxLayout()

        self.buildFileLabel = QLabel("Build File Name - {}(Default)".format(RESOURCES.buildFileNameNoExt))
        self.buildFileInput = QLineEdit()
        self.mergeButton = QPushButton("Merge")
        self.container = QWidget()
        
        self.buildFileInput.setPlaceholderText("Enter Build File Name")

        layout.addWidget(self.buildFileLabel)
        layout.addWidget(self.buildFileInput)
        layout.addWidget(self.mergeButton)
        
        self.container.setLayout(layout)
        self.setCentralWidget(self.container)
        
        self.buildFileInput.textChanged.connect(self.ChangedBuildFileName)
        self.mergeButton.clicked.connect(self.FinalMerge)
    
    def ChangedBuildFileName(self):
        RESOURCES.buildFileNameNoExt = self.buildFileInput.text()
        RESOURCES.buildFileNameExt = AddingFileExt(RESOURCES.buildFileNameNoExt)
        self.buildFileLabel.setText("Build File Name - {}".format(RESOURCES.buildFileNameNoExt))
        

    def FinalMerge(self):
        numExcelToMerge = len(RESOURCES.excelList)
        rawBuildPath = pathlib.PureWindowsPath(os.getcwd())
        pureBuildPath = rawBuildPath.as_posix()
        
        if(numExcelToMerge == 0):
            self.error = ErrorWindow()
            self.error.show()
            ErrorWindow.label.setText("Files to be Merged are not Provided")
        
        elif(self.buildFileNameNoExt.isspace() == True or RESOURCES.buildFileNameNoExt == ""):
            self.error = ErrorWindow()
            self.error.show()
            ErrorWindow.label.setText("Build File Needs to have a Name")
        else:
            
            #Creating build excel
            docBuildWorkbook = Workbook()
            docBuildSheet = docBuildWorkbook.active
            docBuildWorkbook.save(filename=RESOURCES.buildFileNameExt)
        
            #Merge loop
            DATA = []
            numExcelMerged = 0
            headingValue = 0
            
            while(numExcelMerged != numExcelToMerge):
    
                #Opening an excel
                docUserWorkbook = load_workbook(RESOURCES.excelList[numExcelMerged])
                docUserSheet = docUserWorkbook.active


                #Setting row, column parameters
                maxUserRow = docUserSheet.max_row + 1
                maxColumn = docUserSheet.max_column + 1
    
                docBuildSheet.column_dimensions["A"].width = 20
                docBuildSheet.column_dimensions["B"].width = 60
                docBuildSheet.column_dimensions["C"].width = 15
                docBuildSheet.column_dimensions["D"].width = 15
                docBuildSheet.column_dimensions["E"].width = 15
                docBuildSheet.column_dimensions["F"].width = 23
    
    
                #Reading and writing from an excel
                for i in range(1, maxUserRow):
                    userRow = str(i)
                    row = []        
                    for j in range(1, maxColumn):
                        column = chr(64+j)
                        userCell = column + userRow
                        row.append(docUserSheet[userCell].value)
                    row.append(GetExcelFileName(RESOURCES.excelList[numExcelMerged]))
                    
                    if(IfValue(row) == True):

                        if(len(DATA) == 0):
                            DATA.append(row)
                            
                        else:
                            
                            if(type(row[0]) == datetime.datetime):  
                                
                                for i in range (0, len(DATA)):
                                    
                                    if(type(DATA[i][0]) == datetime.datetime):
                                        
                                        if(row[0] < DATA[i][0]):
                                            DATA.insert(i,row)

                            else:
                                DATA.append(row)
                                row[0] = DATETIME(row[0])

                    else:
                        headingValue = headingValue + 1
                        
                        if(headingValue == 1):
                            docBuildSheet.append(row)                
                    
                for i in range (0, len(DATA)):
                    docBuildSheet.append(DATA[i])
                    
                numExcelMerged = numExcelMerged + 1
            RESOURCES.excelList.clear()

            #Saving the Build File
            docBuildWorkbook.save(RESOURCES.buildFileNameExt)
        
            #Clearing Display
            RESOURCES.excelDisplay.clear()
            

            self.AfterMergeWindow = AfterMerge()
            self.AfterMergeWindow.show()

            self.close()