from PySide6.QtWidgets import QVBoxLayout, QWidget, QLabel, QPushButton

class Window_ocorrencias(QWidget):
    def __init__(self, main_window):
        super().__init__()

        self.resize(800, 600)

        self.layout = QVBoxLayout(self)

        self.label_ocorrencias = QLabel("Ocorrências")
        self.layout.addWidget(self.label_ocorrencias)

        self.layout.addWidget(QPushButton("Voltar", clicked=self.return_main_window))

        self.main_window = main_window

    def return_main_window(self):
        self.main_window.return_main_window()