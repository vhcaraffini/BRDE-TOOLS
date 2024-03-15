import win32com.client as win32
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def enviar_email_vencimento_cbt():
    # Caminhos
    root = tk.Tk()
    root.withdraw()
    CAMINHO_EXCEL = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    CAMINHO_IMAGEM = 'C:/Users/e.marcus.machado/OneDrive - Banco Regional de Desenvolvimento do Extremo Sul/Imagens/Assinatura.png'

    # Abrindo aba do Excel
    df1 = pd.read_excel(CAMINHO_EXCEL, sheet_name='Info_Mutuario')
    df2 = pd.read_excel(CAMINHO_EXCEL, sheet_name='Mailing')

    # Gerador
    for i, mutuario_plan1 in enumerate(df1['MUTUARIO']):
        # Definindo formato de envio
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)

        # E-mails de Envio
        for n, mutuario_plan2 in enumerate(df2['MUTUARIO']):
            if mutuario_plan1 == mutuario_plan2:
                email_a_enviar = df2.loc[n, 'E-MAIL']

        email.To = email_a_enviar

        # Informações para o corpo do email
        plano = df1.loc[i, 'PLANO']
        data_vencimento = df1.loc[i, 'DATA_VENCIMENTO'].strftime("%d/%m/%Y")
        valor_parcela = df1.loc[i, 'VALOR_PARCELA']

        # Assunto do e-mail
        email.Subject = f'VENCIMENTO {mutuario_plan1} {data_vencimento}'

        # Corpo do email
        email.BodyFormat = 2
        email.HTMLBody = f"""
        <html>
        <body>
        <p>Prezados, {mutuario_plan1}.</p>
        <p> </p>
        <p> </p>
        <p>Atenciosamente,</p>
        <img src='cid:Assinatura.png'>
        </body>
        </html>
        """
        email.HTMLBody = email.HTMLBody.replace('<body>', '<body><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">')

        # email.SentOnBehalfOfName = 'secob.pr@brde.com.br'

        # Enviando E-mail
        # email.Send()    