from PySide6.QtWidgets import QVBoxLayout, QWidget, QLabel, QPushButton, QLineEdit
from ocorrencias.ocorrencias_email import registrar_ocorrencia_email

class Window_ocorrencias(QWidget):
    def __init__(self, main_window):
        super().__init__()

        # Tamanho da janela em pixels
        self.resize(800, 600)

        self.layout = QVBoxLayout(self)

        # Criando botão ocorrencia
        self.label_ocorrencias = QLabel("Ocorrências")
        self.layout.addWidget(self.label_ocorrencias)

        # Adicionando uma caixa de entrada para a data
        self.date_input = QLineEdit()
        self.date_input.setPlaceholderText("Digite a data (dd/mm/aaaa)")
        self.layout.addWidget(self.date_input)

        # Ações dos botões da janela
        self.layout.addWidget(QPushButton("Voltar", clicked=self.return_main_window))
        self.layout.addWidget(QPushButton("Registrar Ocorrência", clicked=registrar_ocorrencia_email))

        self.main_window = main_window

    # Função que retorna a janela Principal
    def return_main_window(self):
        self.main_window.return_main_window()
        