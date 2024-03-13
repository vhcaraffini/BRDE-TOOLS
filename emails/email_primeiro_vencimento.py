import win32com.client as win32
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def enviar_email_primeiro_vencimento():
    # Caminhos
    CAMINHO_EXCEL = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    CAMINHO_IMAGEM = 'C:/Users/e.marcus.machado/OneDrive - Banco Regional de Desenvolvimento do Extremo Sul/Imagens/Assinatura.png'
    CAMINHO_PDF = 'C:/Users/e.marcus.machado/OneDrive - Banco Regional de Desenvolvimento do Extremo Sul/Documentos/Padrão/BRDE - Internet Banking - Tutorial Acesso.pdf'

    # Abrindo aba do Excel
    df1 = pd.read_excel(CAMINHO_EXCEL, sheet_name='Planilha1')
    df2 = pd.read_excel(CAMINHO_EXCEL, sheet_name='Planilha2')

    # Gerador
    for i, mutuario_plan1 in enumerate(df1['Mutuario']):
        # Definindo formato de envio
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)

        # E-mails de Envio
        for n, mutuario_plan2 in enumerate(df2['Mutuario']):
            if mutuario_plan1 == mutuario_plan2:
                email_a_enviar = df2.loc[n, 'E-MAIL']

        # email.To = email_a_enviar

        # Informações para o corpo do email
        plano = df1.loc[i, 'PLANO']
        data_inicio_carencia = df1.loc[i, 'Data Início Carência'].strftime("%d/%m/%Y")

        # Assunto do e-mail
        email.Subject = f'PRIMEIRO VENCIMENTO {mutuario_plan1} {data_inicio_carencia}'

        # Corpo do email
        email.BodyFormat = 2
        email.HTMLBody = f"""
        <html>
        <body>
        <p>Prezados, {mutuario_plan1}.</p>
        <p> </p>
        <p>Viemos por meio deste e-mail comunicar a data de vencimento da primeira parcela do plano {plano} do mutuário {mutuario_plan1} com o BRDE.</p>
        <p> </p>
        <p>VENCIMENTO: {data_inicio_carencia}.</p>
        <p> </p>
        <p>Enviamos em anexo o tutorial de primeiro acesso à nossa plataforma, onde ficam disponiveis os boletos para pagamento e todas as informações sobre a operação.</p>
        <p> </p>
        <p>Caso tenham alguma dúvida, ficamos à disposição por meio deste e-mail e telefone (41) 3219-8099.</p>
        <p> </p>
        <p>Atenciosamente,</p>
        <img src='cid:Assinatura.png'>
        </body>
        </html>
        """
        email.HTMLBody = email.HTMLBody.replace('<body>', '<body><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">')

        # Adicionando Anexo
        email.Attachments.Add(CAMINHO_PDF)

        email.SentOnBehalfOfName = 'secob.pr@brde.com.br'

        print('Funcionando')
        # Enviando E-mail
        # email.Send()