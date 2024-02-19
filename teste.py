import tkinter as tk
from tkinter import filedialog

def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw() # Esconde a janela principal

    arquivo_selecionado = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))

    if arquivo_selecionado:
        print("Arquivo selecionado:", arquivo_selecionado)
    else:
        print("Nenhum arquivo selecionado.")

selecionar_arquivo()