from PySide6.QtWidgets import QVBoxLayout, QWidget, QLabel, QPushButton


class window_posicao(QWidget):
    def __init__(self, main_window):
        super().__init__()

        self.layout = QVBoxLayout(self)

        self.label_oficios = QLabel("Posição Fim do Mês")
        self.layout.addWidget(self.label_oficios)

        # Conectando o sinal clicked do botão a um método da própria classe
        self.gerar_posicao_fim_mes = QPushButton("Posição Fim do Mês Cresol")
        # self.gerar_posicao_fim_mes.clicked.connect()
        self.layout.addWidget(self.gerar_posicao_fim_mes)

        self.layout.addWidget(QPushButton("Voltar", clicked=self.return_main_window))

        self.main_window = main_window

    def return_main_window(self):
        self.main_window.return_main_window()
