from PySide6.QtWidgets import QVBoxLayout, QWidget, QLabel, QPushButton
from emails.email_primeiro_vencimento import enviar_email_primeiro_vencimento

class window_email(QWidget):
    def __init__(self, main_window):
        super().__init__()

        self.layout = QVBoxLayout(self)

        self.label_oficios = QLabel("Enviador de E-mails")
        self.layout.addWidget(self.label_oficios)

        # Conectando o sinal clicked do botão a um método da própria classe
        self.emai_primeiro_vencimento = QPushButton("Enviar E-mail Primeiro Vencimento")
        self.emai_primeiro_vencimento.clicked.connect(self.enviar_email_primeiro_vencimento_function)
        self.layout.addWidget(self.emai_primeiro_vencimento)

        self.layout.addWidget(QPushButton("Voltar", clicked=self.return_main_window))

        self.main_window = main_window

    def return_main_window(self):
        self.main_window.return_main_window()

    def enviar_email_primeiro_vencimento_function(self):
        enviar_email_primeiro_vencimento()