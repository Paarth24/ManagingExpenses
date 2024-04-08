from PyQt5.QtWidgets import QFileDialog, QWidget
import pathlib
import os
from BuildDirPath import BuildDirPath

class BuildDirWindow(QWidget, BuildDirPath):

    def __init__(self):
        super().__init__()
        
        userPath = QFileDialog.getExistingDirectory(self, "Build Location", "C:/")
        userDefinedPath = userPath
        rawBuildPath = pathlib.PureWindowsPath(userDefinedPath)
        pureBuildPath = str(rawBuildPath.as_posix())
        os.chdir(pureBuildPath)
        BuildDirPath.buildDir.setText(os.getcwd())
        
        self.close()



