from PyQt5.QtWidgets import QLabel
import os

class BuildDirPath():
    def Initialize(self):
        BuildDirPath.buildDir = QLabel(os.getcwd())




