import win32com.client as win32
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def enviar_email_primeiro_vencimento():
    # Caminhos
    CAMINHO_EXCEL = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    CAMINHO_PDF = 'C:/Users/e.marcus.machado/OneDrive - Banco Regional de Desenvolvimento do Extremo Sul/Imagens/BRDE - Internet Banking - Tutorial Acesso.pdf'
    CAMINHO_IMAGEM = 'C:/Users/e.marcus.machado/OneDrive - Banco Regional de Desenvolvimento do Extremo Sul/Imagens/Assinatura.png'

    # Abrindo aba do Excel
    df1 = pd.read_excel(CAMINHO_EXCEL, sheet_name='PRIMEIRO_VENCIMENTO')
    df2 = pd.read_excel(CAMINHO_EXCEL, sheet_name='MAILING')

    # Gerador
    for i, mutuario_plan1 in enumerate(df1['Mutuario']):
        # Definindo formato de envio
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)

        # E-mails de Envio
        for n, mutuario_plan2 in enumerate(df2['Mutuário']):
            if mutuario_plan1 == mutuario_plan2:
                email_a_enviar = df2.loc[n, 'E-MAIL']

        email.To = email_a_enviar

        # Informações para o corpo do email
        plano = df1.loc[i, 'Plano']
        data_inicio_carencia = df1.loc[i, 'Data Início Carência'].strftime("%d/%m/%Y")

        # Assunto do e-mail
        email.Subject = f'PRIMEIRO VENCIMENTO {mutuario_plan1} {data_inicio_carencia}'

        # Corpo do email
        email.BodyFormat = 2
        # Substituindo variáveis no corpo do email
        corpo_email = df1.loc[0, 'CORPO EMAIL HTML'].format(
            MUTUARIO=mutuario_plan1,
            PLANO=plano,
            DATA_VENCIMENTO=data_inicio_carencia,
        )
        
        email.HTMLBody = corpo_email.HTMLBody.replace('<body>', '<body><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">')

        # Adicionando Anexo
        email.Attachments.Add(CAMINHO_PDF)

        email.SentOnBehalfOfName = 'secob.pr@brde.com.br'

        # Enviando E-mail
        email.Send()