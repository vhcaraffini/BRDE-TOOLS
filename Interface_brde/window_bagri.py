from PySide6.QtWidgets import QVBoxLayout, QWidget, QLabel, QPushButton
from bagri.gerador_de_oficio_bagri import gerar_oficio_bagri
from bagri.email_oficio import enviar_oficio_bagri

class window_oficio_bagri(QWidget):
    def __init__(self, main_window):
        super().__init__()

        self.layout = QVBoxLayout(self)

        self.label_oficios = QLabel("Oficios para o Banco do Agricultor")
        self.layout.addWidget(self.label_oficios)

        # Conectando o sinal clicked do botão a um método da própria classe
        self.button_gerar_oficios = QPushButton("Gerar oficios do banco do Agricultor")
        self.button_gerar_oficios.clicked.connect(self.gerador_oficios_bagri)
        self.layout.addWidget(self.button_gerar_oficios)

        self.button_enviar_oficios = QPushButton("Enviar oficios do banco do Agricultor")
        self.button_enviar_oficios.clicked.connect(self.enviador_oficios_bagri)
        self.layout.addWidget(self.button_enviar_oficios)

        self.layout.addWidget(QPushButton("Voltar", clicked=self.return_main_window))

        self.main_window = main_window

    def return_main_window(self):
        self.main_window.return_main_window()

    def gerador_oficios_bagri(self):
        gerar_oficio_bagri()

    def enviador_oficios_bagri(self):
        enviar_oficio_bagri()