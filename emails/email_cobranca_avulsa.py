import win32com.client as win32
from tkinter import filedialog
import tkinter as tk
import pandas as pd
import datetime
import locale

# Definir a localização para português do Brasil
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Obter o nome do mês atual e o numero do ano
nome_mes_atual = datetime.datetime.now().strftime('%B')
ano_atual = datetime.datetime.now().strftime('%Y')

def ponto_para_virgula(valor):
    str_valor_rounded = str(round(valor, 2))
    if '.' in str_valor_rounded:
        novo_valor = ''
        for n in str_valor_rounded:
            if n == '.':
                n = ','
            novo_valor += n
    return novo_valor


def enviar_email_vencimento_cbt():
    # Caminhos
    root = tk.Tk()
    root.withdraw()
    CAMINHO_EXCEL = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    CAMINHO_IMAGEM = 'C:/Users/e.marcus.machado/OneDrive - Banco Regional de Desenvolvimento do Extremo Sul/Imagens/Assinatura.png'

    # Abrindo aba do Excel
    df = pd.read_excel(CAMINHO_EXCEL, sheet_name=f'{nome_mes_atual} - {ano_atual}')

    # Gerador
    for i, mutuario_plan1 in enumerate(df['Mutuário']):
        # Definindo formato de envio
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)

        # E-mails de Envio
        for n, mutuario_plan2 in enumerate(df['MUTUARIO']):
            if mutuario_plan1 == mutuario_plan2:
                email_a_enviar = df.loc[n, 'E-MAIL']

        email.To = email_a_enviar

        # Informações para o corpo do email
        plano = df.loc[i, 'Nº Plano']
        data_vencimento = df.loc[i, 'VENCIMENTO'].strftime("%d/%m/%Y")
        valor_parcela = df.loc[i, 'Valor Prestação R$']

        # Assunto do e-mail
        email.Subject = f'VENCIMENTO {mutuario_plan1} {data_vencimento}'

        # Corpo do email
        email.BodyFormat = 2
        email.HTMLBody = f"""
        <html>
        <body>
        <p>Prezados: {mutuario_plan1}.</p>
        <p>Seguem os valores com o vencimento em {data_vencimento}:</p>
        <p>Plano: {plano} Valor da parcela: R$ {ponto_para_virgula(valor_parcela)}</p>
        <p>Informamos que o documento para pagamento já está disponível no Internet Banking do BRDE.</p>
        <p>Atenciosamente,</p>
        <img src='cid:Assinatura.png'>
        </body>
        </html>
        """
        email.HTMLBody = email.HTMLBody.replace('<body>', '<body><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">')

        email.SentOnBehalfOfName = 'secob.pr@brde.com.br'

        # Enviando E-mail
        # email.Send()