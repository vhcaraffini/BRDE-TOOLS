from PySide6.QtWidgets import QMainWindow, QPushButton, QVBoxLayout, QWidget, QStackedWidget, QLabel
from window_ocorrencias import Window_ocorrencias
from window_bagri import window_oficio_bagri

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("BRDE TOOLS")
        self.resize(800, 600)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout(self.central_widget)

        # Criar um QStackedWidget para gerenciar as diferentes janelas
        self.stacked_widget = QStackedWidget()
        self.layout.addWidget(self.stacked_widget)

        # Botão para voltar à janela principal
        self.button_back = QPushButton("Voltar")
        self.button_back.clicked.connect(self.return_main_window)

        # Janela 1
        self.window_main = QWidget()
        self.layout_window_main = QVBoxLayout(self.window_main)
        self.label_main_window = QLabel("Ferramentas")
        self.layout_window_main.addWidget(self.label_main_window)

        # Criar botões para alternar entre as janelas
        self.button_to_window_ocorrencias = QPushButton("Ocorrências")
        self.layout_window_main.addWidget(self.button_to_window_ocorrencias)
        self.button_to_window_ocorrencias.clicked.connect(self.show_window2)

        self.button_to_window_emails = QPushButton("E-mails")
        self.layout_window_main.addWidget(self.button_to_window_emails)
        self.button_to_window_emails.clicked.connect(self.show_window3)

        self.stacked_widget.addWidget(self.window_main)

        # Janela 2
        self.window_ocorrencias = Window_ocorrencias(self)
        self.stacked_widget.addWidget(self.window_ocorrencias)

        # Janela 3
        self.window_emails = window_oficio_bagri(self)
        self.stacked_widget.addWidget(self.window_emails)

    def return_main_window(self):
        self.stacked_widget.setCurrentIndex(0)
        
    def show_window2(self):
        self.stacked_widget.setCurrentIndex(1)

    def show_window3(self):
        self.stacked_widget.setCurrentIndex(2)
        