from PySide6.QtWidgets import QVBoxLayout, QWidget, QLabel, QPushButton, QLineEdit
from ocorrencias.ocorrencias_email import registrar_ocorrencia_email
from ocorrencias.ocorrencias import registrar_ocorrencia_da_data

class Window_ocorrencias(QWidget):
    def __init__(self, main_window):
        super().__init__()

        self.layout = QVBoxLayout(self)

        # Criando botão ocorrencia
        self.label_ocorrencias = QLabel("Ocorrências")
        self.layout.addWidget(self.label_ocorrencias)

        # Adicionando uma caixa de entrada para a data
        self.data_entrada = QLineEdit()
        self.data_entrada.setPlaceholderText("Data da Ocorrência: (dd/mm/aaaa)")
        self.layout.addWidget(self.data_entrada)

        # Conectar o sinal clicked do botão à função para registrar a ocorrência do referente dia
        registrar_botao_dia = QPushButton("Registrar Ocorrência do referente dia")
        registrar_botao_dia.clicked.connect(self.enviar_data_de_registramento_ocorrencia)
        self.layout.addWidget(registrar_botao_dia)

        registrar_button_email = QPushButton("Registrar Ocorrência de e-mails")
        registrar_button_email.clicked.connect(self.enviar_data_de_registramento_ocorrencia_email)
        self.layout.addWidget(registrar_button_email)

        # Ações dos botões da janela
        self.layout.addWidget(QPushButton("Voltar", clicked=self.return_main_window))

        self.main_window = main_window

    # Função que retorna a janela Principal
    def return_main_window(self):
        self.main_window.return_main_window()

    # Função para registrar a ocorrência do referente dia
    def enviar_data_de_registramento_ocorrencia(self):
        # Obtendo a data atual na caixa de entrada
        date = self.data_entrada.text()

        # Chamando a função de registro de ocorrência do referente dia com a data atual
        registrar_ocorrencia_da_data(date)

    def enviar_data_de_registramento_ocorrencia_email(self):
        # Obtendo a data atual na caixa de entrada
        date = self.data_entrada.text()

        # Chamando a função de registro de ocorrência do referente dia com a data atual
        registrar_ocorrencia_email(date)