from datetime import datetime
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


def data_manuscrita(data):
    # Dicionários de mapeamento para traduzir números e meses para palavras
    meses = {
        1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
        5: "maio", 6: "junho", 7: "julho", 8: "agosto",
        9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
    }

    # Obtendo os componentes da data
    dia = data.day
    mes_numero = data.month
    ano = data.year

    # Obtendo o nome do mês
    mes_nome = meses[mes_numero]

    # Construindo a representação manuscrita da data
    data_manuscrita = f"{dia} de {mes_nome} de {ano}"

    return data_manuscrita


def transforma_ponto_em_virgula(valor):
    numero = ''

    for n in str(valor):
        if n == '.':
            n = ','
        numero += n

    if ',' not in numero:
        numero = numero + ',00'
    if ',' in numero and numero[-2] == ',':
        numero += '0'

    if len(numero) > 6:
        numero = str(numero[:-6]) + '.' + str(numero[-6:])

    return numero

# Obtém os estilos padrão do ReportLab
estilos = getSampleStyleSheet()

# Criação de um estilo para o parágrafo com alinhamento justificado
def configura_paragrafo(n_alignment, tipo, tamanho_font=11, espaco_antes=10, espaco_depois=10, espacamento=14):
    estilo_novo = ParagraphStyle(
    f'estilo_{tipo}',
    parent=estilos['BodyText'],
    alignment=n_alignment,  # 0: Esquerda, 1: Centro, 2: Direita, 4: Justificado
    fontName='Helvetica',  # Substitua pela fonte desejada
    fontSize=tamanho_font,  # Tamanho da fonte
    spaceBefore=espaco_antes,  # Espaço antes do parágrafo
    spaceAfter=espaco_depois,  # Espaço após o parágrafo
    leading=espacamento,  # Espaçamento entre as linhas
)
    return estilo_novo

def cabecalho(canvas, doc, content):
    canvas.saveState()
    w, h = content.wrap(doc.width, doc.topMargin)
    content.drawOn(canvas, 0, doc.height + doc.bottomMargin + doc.topMargin - h)
    canvas.restoreState()

def rodape(canvas, doc, content):
    canvas.saveState()
    w, h = content.wrap(0, 0)
    content.drawOn(canvas, 0, 0)
    canvas.restoreState()

def cabecalho_e_rodape(canvas, doc, header_content, footer_content):
    cabecalho(canvas, doc, header_content)
    rodape(canvas, doc, footer_content)