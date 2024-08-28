from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, PageBreak, PageTemplate, Image
from bagri.funcoes import data_manuscrita, transforma_ponto_em_virgula, configura_paragrafo, cabecalho_e_rodape, get_folder
from reportlab.lib.styles import getSampleStyleSheet
from functions_for_windows import show_popup
from reportlab.platypus.frames import Frame
from datetime import datetime
from reportlab.lib import pagesizes
from reportlab.lib.units import cm
from tkinter import filedialog
from functools import partial
import tkinter as tk
import pandas as pd
import os


def gerar_oficio_bagri():
    CAMINHO_OFICIOS = get_folder()
    PATH_IMAGE = f'{os.path.dirname(os.path.dirname(os.path.realpath(__file__)))}/arquivos'
    
    # Caminho do excel
    root = tk.Tk()
    root.withdraw()

    CAMINHO_ARQUIVO = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    df1 = pd.read_excel(CAMINHO_ARQUIVO, sheet_name='RESUMO')
    df2 = pd.read_excel(CAMINHO_ARQUIVO, sheet_name='REPASSES')

    # Gerando os arquivos baseado nos clientes
    for i, cliente in enumerate(df1['RAZÃO SOCIAL']):
        # Informações planilha RESUMO
        memorando = df1.loc[i, 'Nº MEMO']
        cnpj_cpf = df1.loc[i, 'CNPJ/MF Nº - CPF/MF Nº']
        data_pagamento = df1.loc[i, 'PAGAMENTO'].strftime("%d/%m/%Y")
        data_vencimento = df1.loc[i, 'VENCIMENTO'].strftime("%d/%m/%Y")
        data_repasse = df1.loc[i, 'REPASSE'].strftime("%d/%m/%Y")
        n_convenio = df1.loc[i, 'CONVÊNIO']

        # Informações planilha REPASSES
        convenios_repasse = df2['CODCOOP'].tolist()
        clientes_repasse = df2['NOME CLIENTE'].tolist()

        # Data hoje
        data_manuscrita_hoje = data_manuscrita(datetime.now())

        if n_convenio not in convenios_repasse:
            continue

        if n_convenio == 0 and cliente not in clientes_repasse:
            continue

        s = 0
        if cliente in clientes_repasse:
            for mutuario in clientes_repasse:
                if cliente == mutuario:
                    convenio_avaliador = df2.loc[s, 'CODCOOP']
                s += 1

            if convenio_avaliador != 0:
                continue


        getSampleStyleSheet()

        nome_do_arquivo = F'{CAMINHO_OFICIOS}/Oficio de {cliente}.pdf'

        TAMANHO_DAS_PAGINAS = pagesizes.portrait(pagesizes.A4)

        doc = SimpleDocTemplate(nome_do_arquivo, pagesize=TAMANHO_DAS_PAGINAS, 
            leftMargin = 2.2 * cm, 
            rightMargin = 2.2 * cm,
            topMargin = 3 * cm, 
            bottomMargin = 2 * cm)
        
        frame = Frame(doc.leftMargin, 0, doc.width, doc.height, id='normal')

        conteudo_cabecalho = Image(f'{PATH_IMAGE}/Cabeçalho.jpg', width=21 * cm, height=5 * cm)
        conteudo_rodape = Image(f'{PATH_IMAGE}/Rodapé.jpg', width=21 * cm, height=4 * cm)

        modelo = PageTemplate(id='test', frames=frame, onPage=partial(cabecalho_e_rodape, header_content=conteudo_cabecalho, footer_content=conteudo_rodape))

        doc.addPageTemplates([modelo])

        estilo_justificado = configura_paragrafo(4, 'justificado')
        estilo_a_direita = configura_paragrafo(2, 'direita')
        estilo_centralizado = configura_paragrafo(1, 'centralizado', espacamento=8, espaco_antes=8, espaco_depois=8)
        estilo_assinatura = configura_paragrafo(1, 'assinatura', 9, 6, 6, 6)

        # Adiciona um parágrafo justificado ao PDF
        paragrafos = [
            Paragraph(f"Oficio BRDE/AGCUR Nº {memorando}"),
            Paragraph(f"", estilo_a_direita),        
            Paragraph(f"<b>De:</b><br/>BRDE/AGCUR/GEARC/SECOB"),
            Paragraph(f"", estilo_a_direita),        
            Paragraph(f"<b>Para:</b><br/>{cliente}<br/>CNPJ/MF Nº {cnpj_cpf}"),
            Paragraph(f"", estilo_a_direita),
            Paragraph("<b>Ref.: Reembolso Equalização Programa Banco do Agricultor</b>", estilo_justificado),
            Paragraph(f"", estilo_a_direita),        
            Paragraph("Prezados Senhores:", estilo_justificado),
            Paragraph(f"""Em {data_pagamento} foram pagas por essa Cooperativa parcelas de operações contratadas no âmbito 
            do Programa Banco do Agricultor com vencimento em {data_vencimento}. Após o pagamento, o 
            BRDE enviou o arquivo de solicitação de equalização dessas operações à Fomento Paraná, a 
            qual tendo validado as informações realizou o repasse dos juros a serem equalizados.""", estilo_justificado),
            Paragraph(f"", estilo_a_direita),
            Paragraph(f"No dia {data_repasse} o BRDE efetuou a transferência dos valores constantes na planilha a seguir, diretamente na conta dos mutuários.", estilo_justificado),
            Paragraph(f"", estilo_a_direita),
            Paragraph("Colocamo-nos à disposição para quaisquer esclarecimentos.", estilo_justificado),
            Paragraph(f"", estilo_a_direita),
            Paragraph(f"Curitiba/PR, {data_manuscrita_hoje}.", estilo_a_direita),
            Paragraph(f"", estilo_a_direita),
            Paragraph(f"", estilo_a_direita),
            Paragraph(f"", estilo_a_direita),
            Paragraph(f"", estilo_a_direita),
            Paragraph(f"____________________________________________", estilo_centralizado),
            Paragraph(f"CRISTIANE MEDIANEIRA DE CASTRO COHEN", estilo_centralizado),
            Paragraph(f"CHEFE DE COBRANÇA E TESOURARIA", estilo_assinatura),
            ]

        # Filtro informações da Tabela
        m = 0
        cpfs_conveniados = []
        mutuario_conveniados = []
        repasse_conveniado = []
        valores_repassados_conveniados = []

        for i in convenios_repasse:
            if i == n_convenio:
                cpfs_conveniados.append(df2.loc[m, 'DOC.IDENT.'])
                mutuario_conveniados.append(df2.loc[m, 'NOME CLIENTE'])
                repasse_conveniado.append(df2.loc[m, 'DATA'])
                valores_repassados_conveniados.append(df2.loc[m, 'VALOR R$'])
            m += 1

        # Cabeçalho da tabela
        dados_tabela = [['CPF/MF Nº', 'MUTUÁRIO', 'DATA REPASSE', 'VALOR EM R$']]

        # Colunas e linhas da tabela
        if cliente in mutuario_conveniados:
            c = 0
            for p in mutuario_conveniados:
                if cliente == p:
                    linha = [
                        cpfs_conveniados[c], mutuario_conveniados[c], repasse_conveniado[c].strftime("%d/%m/%Y"), 
                        f'R$ {transforma_ponto_em_virgula(valores_repassados_conveniados[c])}'
                        ]
                    dados_tabela.append(linha)
                c += 1

        if cliente not in mutuario_conveniados and n_convenio != 0:
            for row_count in range(len(mutuario_conveniados)):
                linha = [
                    cpfs_conveniados[row_count], mutuario_conveniados[row_count], repasse_conveniado[row_count].strftime("%d/%m/%Y"), 
                    f'R$ {transforma_ponto_em_virgula(valores_repassados_conveniados[row_count])}'
                    ]
                dados_tabela.append(linha)

        tabela = Table(dados_tabela)

        estilo_tabela = TableStyle([('BACKGROUND', (0, 0), (-1, 0), 'grey'),
                                    ('TEXTCOLOR', (0, 0), (-1, 0), 'white'),
                                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                    ('BACKGROUND', (0, 1), (-1, -1), 'white'),
                                    ('GRID', (0, 0), (-1, -1), 1, 'black')])
        
        tabela.setStyle(estilo_tabela)
        tabela.setStyle(estilo_tabela)

        # Adicionando a proximas paginas
        paragrafos.append(PageBreak())
        paragrafos.append(tabela)

        doc.build(paragrafos)

    show_popup('Oficios Gerados', f"Os oficios foram gerados na pasta:\n\n{CAMINHO_OFICIOS}")