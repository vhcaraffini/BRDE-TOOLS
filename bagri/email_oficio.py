from functions_for_windows import get_mother_folder
import win32com.client as win32
from tkinter import filedialog
import tkinter as tk
import pandas as pd
import os

def enviar_oficio_bagri():
    # Caminhos
    root = tk.Tk()
    root.withdraw()

    CAMINHO_EXCEL = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))    
    CAMINHO_IMAGEM = f'{os.path.dirname(os.path.dirname(os.path.realpath(__file__)))}/arquivos/Assinatura.png'
    GET_PATH = get_mother_folder()

    # Abrindo abas do Excel
    df = pd.read_excel(CAMINHO_EXCEL, sheet_name='RESUMO')

    # Gerador
    for i, cooperativa_mutuario in enumerate(df['CONVÊNIO/MUTUÁRIO']):
        # Definindo formato de envio
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)

        # Verificador de envio
        memorando = str(df.loc[i, 'Nº MEMO'])
        if memorando == 'nan':
            continue

        # E-mails de Envio
        emails = df.loc[i, 'E-MAIL']
        emails_copia = df.loc[i, 'E-MAIL CÓPIA']
        email.To = emails
        email.CC = f'{emails_copia}'

        # Informações para o corpo do email
        data_repasse = df.loc[i, 'REPASSE'].strftime("%d/%m/%Y")
        data_vencimento = df.loc[i, 'VENCIMENTO'].strftime("%d/%m/%Y")

        # Assunto
        email.Subject = f'{cooperativa_mutuario} - REPASSE BANCO DO AGRICULTOR - PARCELA {data_vencimento}'

        # Corpo do email
        email.BodyFormat = 2
        corpo_email = df.loc[0, 'CORPO EMAIL HTML'].format(
            DATA_REPASSE=data_repasse,
            DATA_VENCIMENTO=data_vencimento,
        )
        email.HTMLBody = corpo_email.replace('<body>', '<body><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">')

        # Adicionando Anexo
        convenio_excel = df.loc[i, 'RAZÃO SOCIAL']
        email.Attachments.Add(F'{GET_PATH}/Oficios/Oficio de {convenio_excel}.pdf')

        # Anexar a imagem da assinatura
        attachment = email.Attachments.Add(CAMINHO_IMAGEM)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "minhaassinatura")

        # Enviando em nome de:
        email.SentOnBehalfOfName = 'secob.pr@brde.com.br'

        # Enviando E-mail
        email.Send()