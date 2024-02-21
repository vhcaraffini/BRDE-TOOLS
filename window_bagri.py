from PySide6.QtWidgets import QVBoxLayout, QWidget, QLabel, QPushButton
from bagri.gerador_de_oficio_bagri import gerar_oficio_bagri

class window_oficio_bagri(QWidget):
    def __init__(self, main_window):
        super().__init__()

        self.resize(800, 600)

        self.layout = QVBoxLayout(self)

        self.label_oficios = QLabel("Gerador de Oficios para o Banco do Agricultor")
        self.layout.addWidget(self.label_oficios)

        # Conectando o sinal clicked do botão a um método da própria classe
        self.button_gerar_oficios = QPushButton("Gerar oficios do banco do Agricultor")
        self.button_gerar_oficios.clicked.connect(self.oficios_bagri)
        self.layout.addWidget(self.button_gerar_oficios)

        self.layout.addWidget(QPushButton("Voltar", clicked=self.return_main_window))

        self.main_window = main_window

    def return_main_window(self):
        self.main_window.return_main_window()

    def oficios_bagri(self):
        # Chamar a função gerar_oficio_bagri do módulo gerador_de_oficio_bagri
        gerar_oficio_bagri()