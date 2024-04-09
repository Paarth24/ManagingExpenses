from PyQt5.QtWidgets import QLabel
import os

class BuildDirPath():
    def Initialize(self):

        BuildDirPath.buildDirLabel = QLabel("Build Folder - {}".format(os.getcwd()))




