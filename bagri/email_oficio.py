import win32com.client as win32
import pandas as pd

def enviar_oficio_bagri():
    # Caminhos
    CAMINHO_EXCEL = 'C:/Users/e.marcus.machado/OneDrive - Banco Regional de Desenvolvimento do Extremo Sul/Documentos/OFICIO/OFICIO.xlsx'
    CAMINHO_IMAGEM = 'C:/Users/e.marcus.machado/OneDrive - Banco Regional de Desenvolvimento do Extremo Sul/Imagens/Assinatura.png'
    CAMINHO_PDF = 'C:/Users/e.marcus.machado/OneDrive - Banco Regional de Desenvolvimento do Extremo Sul/Documentos/OFICIO/Oficios'

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
        data_vencimento = df.loc[i, 'REPASSE'].strftime("%d/%m/%Y")

        # Assunto
        email.Subject = f'{cooperativa_mutuario} - REPASSE BANCO DO AGRICULTOR - PARCELA {data_vencimento}'

        # Corpo do email
        email.BodyFormat = 2
        email.HTMLBody = f"""
        <html>
        <body>
        <p>Prezados:</p>
        <p>Segue em anexo comunicado referente ao repasse da equalização de operações Banco do Agricultor referente as parcelas com vencimento em {data_vencimento}.</p>
        <p> </p>
        <p>O BRDE realizou o repasse diretamente na C/C dos mutuários no dia {data_repasse}.</p>
        <p> </p>
        <p>Atenciosamente,</p>
        <img src='cid:Assinatura.png'>
        </body>
        </html>
        """
        email.HTMLBody = email.HTMLBody.replace('<body>', '<body><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">')

        # Adicionando Anexo
        convenio_excel = df.loc[i, 'RAZÃO SOCIAL']
        email.Attachments.Add(F'{CAMINHO_PDF}/Oficio de {convenio_excel}.pdf')

        email.SentOnBehalfOfName = 'secob.pr@brde.com.br'

        # Enviando E-mail
        email.Send()