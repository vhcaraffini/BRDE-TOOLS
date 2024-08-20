import win32com.client as win32
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def ponto_para_virgula(valor):
    str_valor_rounded = str(round(valor, 2))
    if '.' in str_valor_rounded:
        novo_valor = ''
        for n in str_valor_rounded:
            if n == '.':
                n = ','
            novo_valor += n
    return novo_valor

def enviar_email_vencimento_cba():
    # Caminhos
    root = tk.Tk()
    root.withdraw()
    CAMINHO_EXCEL = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    CAMINHO_IMAGEM = 'C:/Users/e.marcus.machado/OneDrive - Banco Regional de Desenvolvimento do Extremo Sul/Imagens/Assinatura.png'

    # Abrindo aba do Excel
    df1 = pd.read_excel(CAMINHO_EXCEL, sheet_name='EMAILS_VALOR')
    df2 = pd.read_excel(CAMINHO_EXCEL, sheet_name='MAILING')

    # Gerador
    for i, mutuario_plan1 in enumerate(df1['Mutuário']):
        # Definindo formato de envio
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)

        # E-mails de Envio
        for n, mutuario_plan2 in enumerate(df2['MUTUARIOS']):
            if mutuario_plan1 == mutuario_plan2:
                email_a_enviar = df2.loc[n, 'E-MAIL']

        email.To = 'e.marcus.machado@brde.com.br'  # email_a_enviar

        # Informações para o corpo do email
        plano = int(df1.loc[i, 'Plano'])
        data_vencimento = df1.loc[i, 'Vencimento'].strftime("%d/%m/%Y")
        valor_parcela = ponto_para_virgula(df1.loc[i, 'Total em R$'])

        # Assunto do e-mail
        email.Subject = f'VENCIMENTO {mutuario_plan1} {data_vencimento}'

        # Corpo do email
        email.BodyFormat = 2

        # Substituindo variáveis no corpo do email
        corpo_email = df1.loc[0, 'CORPO EMAIL HTML'].format(
            PLANO=plano,
            DATA_VENCIMENTO=data_vencimento,
            VALOR_PARCELA=valor_parcela,
            MUTUARIO=mutuario_plan1
        )

        email.HTMLBody = corpo_email.replace('<body>', '<body><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">')

        email.SentOnBehalfOfName = 'secob.pr@brde.com.br'

        # Enviando E-mail
        email.Send()