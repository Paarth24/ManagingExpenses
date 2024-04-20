from openpyxl import Workbook, load_workbook
from PyQt5.QtWidgets import QLabel, QLineEdit, QMainWindow, QPushButton, QVBoxLayout, QWidget

from ErrorWindow import ErrorWindow
from ExcelDataBase import ExcelDataBase
from ExcelDisplayList import ExcelDisplayList
from Functions import GetExcelFileName, AddingFileExt
from MergeConfirm import MergeConfirm




class MergeWindow(ErrorWindow, ExcelDataBase, ExcelDisplayList):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Merge Window")
        self.setMinimumSize(250, 100)
        
        self.buildFileNameNoExt = "Final_Bank_Statement"
        self.buildFileNameExt = AddingFileExt(self.buildFileNameNoExt)

        layout = QVBoxLayout()

        self.buildFileLabel = QLabel("Build File Name - {}(Default)".format(self.buildFileNameNoExt))
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
        self.buildFileNameNoExt = self.buildFileInput.text()
        self.buildFileNameExt = AddingFileExt(self.buildFileNameNoExt)
        self.buildFileLabel.setText("Build File Name - {}".format(self.buildFileNameNoExt))
        

    def FinalMerge(self):
        numExcelToMerge = len(ExcelDataBase.excelList)
        
        if(numExcelToMerge == 0):
            self.error = ErrorWindow()
            self.error.show()
            ErrorWindow.label.setText("Files to be Merged are not Provided")
        
        elif(self.buildFileNameNoExt.isspace() == True or self.buildFileNameNoExt == ""):
            self.error = ErrorWindow()
            self.error.show()
            ErrorWindow.label.setText("Build File Needs to have a Name")
        
        else:
            
            #Creating build excel
            docBuildWorkbook = Workbook()
            docBuildSheet = docBuildWorkbook.active
            docBuildWorkbook.save(filename=self.buildFileNameExt)
        
            #Merge loop
            numExcelMerged = 0
            while(numExcelMerged != numExcelToMerge):
    
                #Opening an excel
                docUserWorkbook = load_workbook(ExcelDataBase.excelList[numExcelMerged])
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
                    row.append(GetExcelFileName(self.excelList[numExcelMerged]))
    
                numExcelMerged = numExcelMerged + 1
        
            ExcelDataBase.excelList.clear()

            #Saving the Build File
            docBuildWorkbook.save(self.buildFileNameExt)
        
            #Clearing Display
            ExcelDisplayList.excelDisplay.clear()
            

            self.MergeConfirmWindow = MergeConfirm()
            self.MergeConfirmWindow.show()

            self.close()
            

   

