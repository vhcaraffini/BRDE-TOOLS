from functions_for_windows import show_popup
from tkinter import filedialog
import win32com.client as win32
import tkinter as tk
import pandas as pd
import os

def enviar_email_extrato():
    # Caminhos
    # Abrindo Excel
    root = tk.Tk()
    root.withdraw()
    CAMINHO_EXCEL = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    CAMINHO_IMAGEM = f'{os.path.dirname(os.path.dirname(os.path.realpath(__file__)))}/arquivos/Assinatura.png'

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
        email.CC = f'{email_copia}; e.beatriz.juliatto@brde.com.br'

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
        
        # Anexar a imagem da assinatura
        attachment = email.Attachments.Add(CAMINHO_IMAGEM)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "minhaassinatura")

        # Enviando em nome de:
        email.SentOnBehalfOfName = 'acompanhamento.pr@brde.com.br'

        # Enviando E-mail
        email.Send()

    show_popup('E-mails', f"E-mails de Extrato enviados com sucesso")