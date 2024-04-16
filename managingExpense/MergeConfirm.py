from PyQt5.QtCore import QSize
from PyQt5.QtGui import QImage
from PyQt5.QtWidgets import QAbstractButton, QMainWindow, QLabel, QPushButton, QVBoxLayout, QWidget

class MergeConfirm(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Merge Confirmation")
        self.setFixedSize(QSize(150, 100))

        layout = QVBoxLayout()
        container = QWidget()
        label = QLabel("Merged")
        openFileButton = QPushButton("Open File")

        layout.addWidget(label)
        layout.addWidget(openFileButton)
        container.setLayout(layout)
        self.setCentralWidget(container)

            
        
        




