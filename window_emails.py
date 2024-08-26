from PySide6.QtWidgets import QVBoxLayout, QWidget, QLabel, QPushButton
from emails.email_primeiro_vencimento import enviar_email_primeiro_vencimento
from emails.email_vencimento import enviar_email_vencimento_cba
from emails.email_cobranca_avulsa import enviar_email_cobranca_avulsa
from emails.email_extrato import enviar_email_extrato

class window_email(QWidget):
    def __init__(self, main_window):
        super().__init__()

        self.layout = QVBoxLayout(self)

        self.label_oficios = QLabel("Enviador de E-mails")
        self.layout.addWidget(self.label_oficios)

        # Conectando o sinal clicked do botão a um método da própria classe
        self.email_vencimento_cba = QPushButton("Enviar E-mail Parcela")
        self.email_vencimento_cba.clicked.connect(self.enviador_mail_vencimento_cba)
        self.layout.addWidget(self.email_vencimento_cba)

        self.email_cobranca_avulsa = QPushButton("Enviar E-mail Cobrança Avulsa")
        self.email_cobranca_avulsa.clicked.connect(self.enviador_mail_cobranca_avulsa)
        self.layout.addWidget(self.email_cobranca_avulsa)

        self.email_primeiro_vencimento = QPushButton("Enviar E-mail Primeiro Vencimento")
        self.email_primeiro_vencimento.clicked.connect(self.enviador_email_primeiro_vencimento)
        self.layout.addWidget(self.email_primeiro_vencimento)

        self.email_extrato = QPushButton("Enviar E-mail Extrato")
        self.email_extrato.clicked.connect(self.enviador_mail_extrato)
        self.layout.addWidget(self.email_extrato)

        self.layout.addWidget(QPushButton("Voltar", clicked=self.return_main_window))

        self.main_window = main_window

    def return_main_window(self):
        self.main_window.return_main_window()

    def enviador_email_primeiro_vencimento(self):
        enviar_email_primeiro_vencimento()

    def enviador_mail_vencimento_cba(self):
        enviar_email_vencimento_cba()

    def enviador_mail_cobranca_avulsa(self):
        enviar_email_cobranca_avulsa()

    def enviador_mail_extrato(self):
        enviar_email_extrato()