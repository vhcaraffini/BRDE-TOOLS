import win32com.client as win32
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def enviar_email_extrato():
    # Caminhos
    # Abrindo Excel
    root = tk.Tk()
    root.withdraw()
    CAMINHO_EXCEL = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    CAMINHO_IMAGEM = 'C:/Users/e.marcus.machado/OneDrive - Banco Regional de Desenvolvimento do Extremo Sul/Imagens/Assinatura.png'

    # Abrindo aba do Excel
    df1 = pd.read_excel(CAMINHO_EXCEL, sheet_name='EMAIL_EXTRATO')
    df2 = pd.read_excel(CAMINHO_EXCEL, sheet_name='MAILING')

    # Gerador
    for i, empresa in enumerate(df1['EMPRESA']):
        # Definindo formato de envio
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)

        # Informações para o corpo do email
        contrato = df1.loc[i, 'CONTRATO']
        conta = df1.loc[i, 'CONTA']
        email_a_enviar = df1.loc[i, 'PARA']
        email_copia = df1.loc[i, 'CC']
        
        # Destinatarios
        email.To = email_a_enviar
        email.CC = f'{email_copia}; e.beatriz.juliatto@brde.com.br; acompanhamento.pr@brde.com.br'

        # Assunto do e-mail
        email.Subject = f'Solicitação Extrato Atualizado - {empresa}'

        # Corpo do email
        email.BodyFormat = 2
        corpo_email = df1.loc[0, 'CORPO EMAIL HTML'].format(
            EMPRESA=empresa,
            CONTRATO=contrato,
            CONTA=conta,
        )

        email.HTMLBody = corpo_email.replace('<body>', '<body><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">')

        # Remetente
        email.SentOnBehalfOfName = 'e.marcus.machado@brde.com.br'

        # Enviando E-mail
        email.Send()