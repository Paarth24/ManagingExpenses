from openpyxl import Workbook, load_workbook
import os
import pathlib
from PyQt5.QtCore import QSize
from PyQt5.QtWidgets import QApplication, QBoxLayout, QFileDialog, QGridLayout, QHBoxLayout, QLabel, QLineEdit, QMainWindow, QPushButton, QVBoxLayout, QWidget



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
        self.excelList = []
        
        mainLayout = QHBoxLayout()
#-------------------Initialization--------------------------- 
#-------------------Formatting---------------------------         
        mainLayout.addWidget(mergeButton)
        mainLayout.addWidget(addButton)
        mainLayout.addWidget(buildButton)

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
            self.excelList.append(excelName[0])
            print(self.excelList)
        
    
    def ChangingBuildDirectory(self):
        #Changing build location
        userDefinedPath = r"C:\Paarth\ManagingBankStatement\managingExpense\Build"    
        rawBuildPath = pathlib.PureWindowsPath(userDefinedPath)
        pureBuildPath = str(rawBuildPath.as_posix())
        os.chdir(pureBuildPath)
        
    def MergingAddedFiles(self):
        numExcelToMerge = len(self.excelList)
        
        #Creating build excel
        docBuildWorkbook = Workbook()
        docBuildSheet = docBuildWorkbook.active
        docBuildWorkbook.save(filename="Final_Bank_Statement.xlsx")
        
        #Merge loop
        numExcelMerged = 0
        while(numExcelMerged != numExcelToMerge):
    
            #Opening an excel
            docUserWorkbook = load_workbook(self.excelList[numExcelMerged])
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
    
            docBuildSheet.append(["Above is the excel - {}".format(self.excelList[numExcelMerged])])
            docBuildSheet.append([])
    
            numExcelMerged = numExcelMerged + 1
        #Saving the Build File
        docBuildWorkbook.save("Final_Bank_Statement.xlsx")
        print("Merged")
#-------------------Functions---------------------------          


if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()

app.exec_()
