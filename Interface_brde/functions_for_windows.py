from PySide6.QtWidgets import QMessageBox, QApplication, QSpacerItem, QSizePolicy
from PySide6.QtGui import QIcon
import os

diretorio_atual = os.path.dirname(os.path.realpath(__file__))

def show_popup(titulo, texto):
    # Cria uma instância de QMessageBox
    msg = QMessageBox()

    # Configura o texto e o título do pop-up
    msg.setWindowTitle(titulo)
    msg.setText(texto)

    # Define os botões do pop-up
    msg.setStandardButtons(QMessageBox.Ok)

    # Aplicar estilo personalizado
    msg.setStyleSheet("""
        QMessageBox {
            background-color: #F5F5DC;
            border: 5px 5px 5px 5px solid back;
        }
        QLabel {
            font-size: 16px;
            font-family: Candara;
            width: 75px;
            margin: 10px 40px 10px 40px;
        }

        QPushButton {
            background-color: #007553;
            color: #FFFFFF;
            border-radius: 5px;
            padding: 10px;
            font-size: 14px;
            font-family: Candara;
            width: 25px;
        }
        QPushButton:hover {
            background-color: #009688;
        }
    """)

    # Definir a largura e altura da caixa
    msg.setFixedSize(600, 150)

    # Define o Icone
    icon_path = f"{diretorio_atual}/arquivos/BRDE_favicon.png"
    msg.setWindowIcon(QIcon(icon_path))

    # Exibe o pop-up e captura a resposta do usuário
    msg.exec()