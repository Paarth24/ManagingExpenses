from PyQt5.QtCore import QSize
from openpyxl import Workbook, load_workbook
from PyQt5.QtWidgets import QLabel, QLineEdit, QPushButton, QVBoxLayout, QWidget

from GlobalVar import GlobalVar

class MergeWindow(QWidget):
    def __init__(self):
        super().__init__()
        
        self.buildFileName = "Final_Bank_Statement.xlsx"

        layout = QVBoxLayout()
        self.buildFileLabel = QLabel("Final file name - {}".format(self.buildFileName))
        self.buildFileInput = QLineEdit()
        self.mergeButton = QPushButton("Merge")

        layout.addWidget(self.buildFileLabel)
        layout.addWidget(self.buildFileInput)
        layout.addWidget(self.mergeButton)
        self.setLayout(layout)
        
        self.mergeButton.clicked.connect(self.FinalMerge)
        
    def FinalMerge(self):
        numExcelToMerge = len(GlobalVar().ExcelList())
        
        #Creating build excel
        docBuildWorkbook = Workbook()
        docBuildSheet = docBuildWorkbook.active
        docBuildWorkbook.save(filename="Final_Bank_Statement.xlsx")
        
        #Merge loop
        numExcelMerged = 0
        while(numExcelMerged != numExcelToMerge):
    
            #Opening an excel
            docUserWorkbook = load_workbook(GlobalVar().ExcelList()[numExcelMerged])
            docUserSheet = docUserWorkbook.active


            #Setting row, column parameters
            maxUserRow = docUserSheet.max_row + 1
            maxColumn = docUserSheet.max_column + 1
    
            docBuildSheet.column_dimensions["A"].width = 20
            docBuildSheet.column_dimensions["B"].width = 60
            docBuildSheet.column_dimensions["C"].width = 15
            docBuildSheet.column_dimensions["D"].width = 15
            docBuildSheet.column_dimensions["E"].width = 15
    
    
            #Reading and writing from an excel
            for i in range(1, maxUserRow):
                userRow = str(i)
                buildRow = str(i)
                row = []        
                for j in range(1, maxColumn):
                    column = chr(64+j)
                    userCell = column + userRow
                    buildCell = column + buildRow
                    row.append(docUserSheet[userCell].value)            
        
                docBuildSheet.append(row)
    
            docBuildSheet.append(["Above is the excel - {}".format(GlobalVar().ExcelList()[numExcelMerged])])
            docBuildSheet.append([])
    
            numExcelMerged = numExcelMerged + 1
        #Saving the Build File
        docBuildWorkbook.save("Final_Bank_Statement.xlsx")
        print("Merged")



