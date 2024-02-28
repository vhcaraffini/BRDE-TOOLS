from PySide6.QtWidgets import QVBoxLayout, QWidget, QLabel, QPushButton
from termos.separador_de_termos import separa_termos

class window_termos(QWidget):
    def __init__(self, main_window):
        super().__init__()

        self.resize(800, 600)

        self.layout = QVBoxLayout(self)

        self.label_oficios = QLabel("Gestor de Termos")
        self.layout.addWidget(self.label_oficios)

        # Conectando o sinal clicked do botão a um método da própria classe
        self.separar_termos = QPushButton("Separar Termos")
        self.separar_termos.clicked.connect(self.separador_de_termos)
        self.layout.addWidget(self.separar_termos)

        self.layout.addWidget(QPushButton("Voltar", clicked=self.return_main_window))

        self.main_window = main_window

    def return_main_window(self):
        self.main_window.return_main_window()

    def separador_de_termos(self):
        separa_termos()