from BuildDirPath import BuildDirPath
from ExcelDisplayList import ExcelDisplayList
from MainWindow import MainWindow
from ExcelDataBase import ExcelDataBase
from PyQt5.QtWidgets import QApplication
        
if __name__ == "__main__":
    app = QApplication([])
    
    ExcelDataBase().Initialize()
    ExcelDisplayList().Initialize()
    BuildDirPath().Initialize()
    
    window = MainWindow()
    
    window.show()

app.exec_()
