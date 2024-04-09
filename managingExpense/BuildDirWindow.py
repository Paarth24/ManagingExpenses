from PyQt5.QtGui import QCloseEvent, QWindow
from PyQt5.QtWidgets import QFileDialog, QWidget
import pathlib
import os
from BuildDirPath import BuildDirPath

class BuildDirWindow(BuildDirPath):

    def __init__(self):
        super().__init__()
        
        userPath = QFileDialog.getExistingDirectory(self, "Build Location", "C:/")
        userDefinedPath = userPath
        rawBuildPath = pathlib.PureWindowsPath(userDefinedPath)
        pureBuildPath = str(rawBuildPath.as_posix())
        os.chdir(pureBuildPath)
        BuildDirPath.buildDirLabel.setText(os.getcwd())
        
    def closeEvent(self, event):
       self.close()




