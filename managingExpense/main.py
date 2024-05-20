from Resources import RESOURCES
from MainWindow import MainWindow
from PyQt5.QtWidgets import QApplication
from MergeWindow import MergeWindow
        
if __name__ == "__main__":
    app = QApplication([])
    
    RESOURCES().Initialize()
    
    window = MainWindow()
    
    window.show()

app.exec_()

#pushing to git
