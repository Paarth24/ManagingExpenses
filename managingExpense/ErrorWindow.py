from PyQt5.QtCore import QSize
from PyQt5.QtWidgets import QLabel, QMainWindow, QVBoxLayout, QWidget


class ErrorWindow(QMainWindow, QWidget):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Error")
        self.setFixedSize(QSize(250, 100))
        
        layout = QVBoxLayout()
        container = QWidget()
        ErrorWindow.label = QLabel()
        
        layout.addWidget(ErrorWindow.label)
        container.setLayout(layout)
        
        self.setCentralWidget(container)
        




