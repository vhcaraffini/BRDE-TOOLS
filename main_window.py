from PySide6.QtWidgets import QMainWindow, QPushButton, QVBoxLayout, QWidget, QStackedWidget, QLabel
from window_ocorrencias import Window_ocorrencias
from window_bagri import window_oficio_bagri
from window_termos import window_termos

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("BRDE TOOLS")
        self.resize(700, 500)

        # Configuração global de estilo
        self.setStyleSheet("""
            QMainWindow {
                background-color: #8FBC8F;
                border: 0 5px 5px 5px solid back;
                
            }
            QPushButton {
                background-color: #006400;
                border: 4px 4px 4px 4px solid black;
                border-radius: 5px;
                padding: 10px;
                margin: 0px 40px 10px 40px;
                color: #FFFFFF;
                width: 75px;
                height: 40px;
                font-family: Roboto;
                font-size: 17px;
            }
            QLabel {
                color: #000000;
                width: 100px;
                height: 50px;
                margin: 0px 70px 0 70px;
                padding: 0;
                border: 2px 2px 2px 2px solid black;
                font-family: Roboto;
                font-size: 26px;
            }
        """)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout(self.central_widget)

        # Criar um QStackedWidget para gerenciar as diferentes janelas
        self.stacked_widget = QStackedWidget()
        self.layout.addWidget(self.stacked_widget)

        # Adicione as outras janelas aqui
        self.add_windows()

    def add_windows(self):
        # Janela principal
        self.window_BRDE_TOOLS = QWidget()
        layout_window_main = QVBoxLayout(self.window_BRDE_TOOLS)
        label_main_window = QLabel("BRDE TOOLS")
        layout_window_main.addWidget(label_main_window)

        # Botões para alternar entre as janelas
        button_to_window_ocorrencias = QPushButton("Ocorrências")
        button_to_window_ocorrencias.clicked.connect(self.show_window2)
        layout_window_main.addWidget(button_to_window_ocorrencias)

        button_to_window_bagri = QPushButton("Oficios Banco do Agricultor")
        button_to_window_bagri.clicked.connect(self.show_window3)
        layout_window_main.addWidget(button_to_window_bagri)

        button_to_window_termos = QPushButton("Separador de Termos")
        button_to_window_termos.clicked.connect(self.show_window4)
        layout_window_main.addWidget(button_to_window_termos)

        self.stacked_widget.addWidget(self.window_BRDE_TOOLS)

        # Janelas específicas
        self.window_ocorrencias = Window_ocorrencias(self)
        self.stacked_widget.addWidget(self.window_ocorrencias)

        self.window_bagri = window_oficio_bagri(self)
        self.stacked_widget.addWidget(self.window_bagri)

        self.window_termo = window_termos(self)
        self.stacked_widget.addWidget(self.window_termo)

    def return_main_window(self):
        self.stacked_widget.setCurrentIndex(0)

    def show_window2(self):
        self.stacked_widget.setCurrentIndex(1)

    def show_window3(self):
        self.stacked_widget.setCurrentIndex(2)

    def show_window4(self):
        self.stacked_widget.setCurrentIndex(3)