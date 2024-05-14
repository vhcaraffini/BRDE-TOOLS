from PySide6.QtWidgets import QVBoxLayout, QWidget, QLabel, QPushButton, QCalendarWidget, QHBoxLayout, QBoxLayout
from ocorrencias.ocorrencias_email import registrar_ocorrencia_email
from ocorrencias.ocorrencias import registrar_ocorrencia_da_data

class Window_ocorrencias(QWidget):
    def __init__(self, main_window):
        super().__init__()

        # Criando um QVBoxLayout
        self.layout = QVBoxLayout(self)

        # Criando um QHBoxLayout
        layout_horizontal = QHBoxLayout()

        # Criando botão ocorrencia
        self.label_ocorrencias = QLabel("\n\n\n          Ocorrências")
        layout_horizontal.addWidget(self.label_ocorrencias)

        # Adicionando QCalendarWidget para selecionar a data
        self.calendar_widget = QCalendarWidget()
        self.calendar_widget.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self.calendar_widget.setStyleSheet("""
            QCalendarWidget QToolButton {
            background-color: #1e90ff;
            color: green;
        }
            QCalendarWidget QToolButton:hover {
            background-color: #4682b4;
        }
        """)
        self.calendar_widget.setMaximumHeight(100)
        self.calendar_widget.setMinimumHeight(250)
        self.calendar_widget.setMaximumWidth(300)
        self.calendar_widget.setMinimumWidth(100)
        
        # Adicionando o QCalendarWidget ao QHBoxLayout
        layout_horizontal.addWidget(self.calendar_widget)

        # Adicionando o QHBoxLayout ao QVBoxLayout
        self.layout.addLayout(layout_horizontal)

        self.label_espacos = QLabel("")
        self.layout.addWidget(self.label_espacos)
    
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
        # Obtendo a data selecionada no calendário
        selected_date = self.calendar_widget.selectedDate()

        # Convertendo a data para uma string no formato dd/MM/yyyy
        date = selected_date.toString("dd/MM/yyyy")

        # Chamando a função de registro de ocorrência do referente dia com a data selecionada
        registrar_ocorrencia_da_data(date)

    def enviar_data_de_registramento_ocorrencia_email(self):
        # Obtendo a data selecionada no calendário
        selected_date = self.calendar_widget.selectedDate()

        # Convertendo a data para uma string no formato dd/MM/yyyy
        date = selected_date.toString("dd/MM/yyyy")

        # Chamando a função de registro de ocorrência do referente dia com a data selecionada
        registrar_ocorrencia_email(date)