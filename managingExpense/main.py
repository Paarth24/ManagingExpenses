from fileinput import filename
from turtle import width
import openpyxl
from openpyxl import Workbook, load_workbook
import os
import pathlib

#Creating build excel
docBuildWorkbook = Workbook()
docBuildSheet = docBuildWorkbook.active
docBuildWorkbook.save(filename="Final_Bank_Statement.xlsx")

#Number of excels to be merged
excelList = ["example_statement1.xlsx","example_statement2.xlsx","example_statement3.xlsx"]
numExcelToMerge = len(excelList)


#Merge loop
numExcelMerged = 0
while(numExcelMerged != numExcelToMerge):
    
    #Opening an excel
    docUserWorkbook = load_workbook(excelList[numExcelMerged])
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
    
    docBuildSheet.append(["Above is the excel - {}".format(excelList[numExcelMerged])])
    docBuildSheet.append([])
    
    numExcelMerged = numExcelMerged + 1
    
    
#Changing build location
userDefinedPath = r"C:\Paarth\ManagingBankStatement\managingExpense\Build"    
rawBuildPath = pathlib.PureWindowsPath(userDefinedPath)
pureBuildPath = str(rawBuildPath.as_posix())
os.chdir(pureBuildPath)

#Saving the Build File
docBuildWorkbook.save("Final_Bank_Statement.xlsx")




