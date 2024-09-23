from functions_for_windows import show_popup
import win32com.client as win32
from tkinter import filedialog
import tkinter as tk
import pandas as pd
import datetime
import locale
import os

# Definir a localização para português do Brasil
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Obter o nome do mês atual e o numero do ano
nome_mes_atual = datetime.datetime.now().strftime('%B')
ano_atual = datetime.datetime.now().strftime('%Y')
mes_atual = datetime.datetime.now().month

datas_habituais = [f'01/{mes_atual}/{ano_atual}', f'10/{mes_atual}/{ano_atual}', 
                   f'15/{mes_atual}/{ano_atual}', f'18/{mes_atual}/{ano_atual}',
                   f'23/{mes_atual}/{ano_atual}', f'20/{mes_atual}/{ano_atual}',]

# Função para formatar o número como moeda brasileira
def ponto_para_virgula(valor):
    try:
        # Formatar como moeda com separador de milhar e duas casas decimais
        valor_formatado = locale.currency(valor, grouping=True, symbol=False)
        # Remover o símbolo da moeda (R$), deixando só o valor
        return valor_formatado.replace('R$', '').strip()
    except Exception as e:
        print(f"Erro ao formatar valor: {str(e)}")
        return str(valor)


def enviar_email_cobranca_avulsa():
    # Caminhos
    root = tk.Tk()
    root.withdraw()
    CAMINHO_EXCEL = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    CAMINHO_IMAGEM = f'{os.path.dirname(os.path.dirname(os.path.realpath(__file__)))}/arquivos/Assinatura.png'

    # Abrindo aba do Excel
    df1 = pd.read_excel(CAMINHO_EXCEL, sheet_name='COBRANÇA_AVULSA')
    df2 = pd.read_excel(CAMINHO_EXCEL, sheet_name=f'MAILING')

    # Gerador
    for i, mutuario_plan1 in enumerate(df1['Mutuário']):
        # Informações para o corpo do email
        plano = df1.loc[i, 'Plano']
        data_vencimento = df1.loc[i, 'Vencimento'].strftime("%d/%m/%Y")
        valor_parcela = ponto_para_virgula(df1.loc[i, 'Total em R$'])

        # Definindo formato de envio
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)

        if data_vencimento in datas_habituais:
            continue

        # E-mails de Envio
        for n, mutuario_plan2 in enumerate(df2['MUTUARIOS']):
            if mutuario_plan1 == mutuario_plan2:
                email_a_enviar = df2.loc[n, 'E-MAIL']

        email.To = email_a_enviar

        # Assunto do e-mail
        email.Subject = f'VENCIMENTO {mutuario_plan1} {data_vencimento}'

        # Corpo do email
corpo_email = df1.loc[0, 'CORPO EMAIL HTML'].format(
    MUTUARIO=mutuario_plan1,
    PLANO=plano,
    DATA_VENCIMENTO=data_vencimento,
    VALOR_PARCELA=ponto_para_virgula(df1.loc[i, 'Total em R$']),
)
        )
        email.HTMLBody = corpo_email.replace('<body>', '<body><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">')

        # Anexar a imagem da assinatura
        attachment = email.Attachments.Add(CAMINHO_IMAGEM)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "minhaassinatura")

        # Enviando em nome de:
        email.SentOnBehalfOfName = 'secob.pr@brde.com.br'

        # Enviando E-mail
        email.Send()
    
    show_popup('E-mails', f"E-mails enviados com sucesso")
