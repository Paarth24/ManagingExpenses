from PyQt5.QtWidgets import QListWidget, QLabel
import os

class RESOURCES():
    def Initialize(self):
        
        RESOURCES.excelList = []
        RESOURCES.excelDisplay = QListWidget()
        RESOURCES.buildDirLabel = QLabel("Build Folder - {}".format(os.getcwd()))
        RESOURCES.buildFileNameNoExt = ""
        RESOURCES.buildFileNameExt = ""




