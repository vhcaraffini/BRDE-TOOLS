import PyPDF2
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

def separa_termos():
    # CAMINHOS ARQUIVOS
    root = tk.Tk()
    root.withdraw() # Esconde a janela principal

    CAMINHO_EXCEL = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    CAMINHO_ENTRADA_PDF = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do PDF", "*.pdf"), ("Todos os arquivos", "*.*")))

    CAMINHO = 'C:/Users/e.marcus.machado/OneDrive - Banco Regional de Desenvolvimento do Extremo Sul/Documentos/Termos de Quitação/'
    CAMINHO_SAIDA_PDF = f'{CAMINHO}Gerados'

    df1 = pd.read_excel(CAMINHO_EXCEL, sheet_name='Planilha1')

    def pdf_para_paginas(entrada_do_arquivo, saida_da_pasta):
        with open(entrada_do_arquivo, 'rb') as arquivo:
            leitor_pdf = PyPDF2.PdfReader(arquivo)

            # Cria a pasta de saída se não existir
            os.makedirs(saida_da_pasta, exist_ok=True)

            # Itera sobre as páginas e salva cada uma como um novo arquivo PDF
            for numero_pagina, pagina in enumerate(leitor_pdf.pages, start=1):
                nome_mutuario = df1.loc[numero_pagina - 1, 'MUTUARIO']
                escritor_pdf = PyPDF2.PdfWriter()
                escritor_pdf.add_page(pagina)

                saida_do_arquivo = os.path.join(saida_da_pasta, f'{nome_mutuario}.pdf')
                with open(saida_do_arquivo, 'wb') as saida:
                    escritor_pdf.write(saida)


    # Certifica-se de que a pasta de saída exista ou cria uma nova
    os.makedirs(CAMINHO_SAIDA_PDF, exist_ok=True)

    pdf_para_paginas(CAMINHO_ENTRADA_PDF, CAMINHO_SAIDA_PDF)