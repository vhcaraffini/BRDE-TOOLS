from tkinter import filedialog
import tkinter as tk
import pandas as pd
import openpyxl

def data_planilha(CAMINHO_ARQUIVO):
    data = ''
    for n in range(10):
        data = CAMINHO_ARQUIVO[-1-n] + data
    return data


def posicao_coperativas_credicopavel():
    # Abrindo Excel
    root = tk.Tk()
    root.withdraw()

    CAMINHO_ARQUIVO = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    df = pd.read_excel(CAMINHO_ARQUIVO, sheet_name='Planilha1')
    CAMINHO_NOVO = F'M:/Secob/POSIÇÃO FINAL MÊS/CREDICOOPAVEL - POSIÇÃO {data_planilha(CAMINHO_ARQUIVO)}'

    filtro = df['CodInt'] = 1007 & df['CodInt'].notna()
    colunas_desejadas = ['Mutuario', 'CGC', 'PLANO', 'Vencimento Mais Antigo', 'NR', 'Saldo Vencido', 'Saldo Total', 'CodInt', 'CodCoop']
    df_filtrado = df[filtro]
    df_filtrado_coluas = df_filtrado[colunas_desejadas]
    tabela_excel = df_filtrado_coluas.to_excel(CAMINHO_NOVO, index=False, na_rep='')

    wb = openpyxl.load_workbook(CAMINHO_NOVO)
    ws = wb.active

    # Iterar sobre as colunas e ajustar a largura
    for coluna in ws.columns:
        max_length = 0
        column = [col.value for col in coluna]
        for cell in column:
            try:
                if len(str(cell)) > max_length:
                    max_length = len(cell)
            except:
                pass
        adjusted_width = (max_length + 4)
        ws.column_dimensions[coluna[0].column_letter].width = adjusted_width
    
    # Salvar as alterações
    wb.save(CAMINHO_NOVO)