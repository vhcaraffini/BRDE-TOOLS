from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from datetime import datetime, timedelta 
from selenium import webdriver
from tkinter import filedialog
from time import sleep
import tkinter as tk
import openpyxl

TIMER = 1


def registrar_ocorrencia_email(data):
    # Dia do registro
    dia_formatado_preencher = data

    # Abrindo Excel
    root = tk.Tk()
    root.withdraw() # Esconde a janela principal

    CAMINHO_ARQUIVO = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    workbook = openpyxl.load_workbook(CAMINHO_ARQUIVO)
    sheet = workbook['Planilha1']

    # Valores
    mutuarios = []
    descricoes = []
    cobranca = '13'
    mutuarios_anteriores = []

    # Pegando valores do excel e adicionando a lista
    for row in sheet.iter_rows(values_only=True):
        mutuarios.append(row[0])
        descricoes.append(row[1])

    for n in range(len(mutuarios)):
        mutuario = mutuarios[n]
        descricao = descricoes[n]
        nome = 'Cris'

        # Filtrando nomes iguais
        if mutuario in mutuarios_anteriores:
            continue

        # Mostrando Mutuario
        print(mutuario)

        # Acessando o site
        driver = webdriver.Chrome()
        driver.get(f"https://brbank.brde.com.br/Pessoas/Buscar")

        # Encontrando e preenchendo barra de pesquisa
        encontrando_barra_pesquisa = WebDriverWait(driver, TIMER).until(EC.presence_of_element_located((By.ID, 'NomeCnpjCpf')))
        encontrando_barra_pesquisa.send_keys(mutuario)
        encontrando_barra_pesquisa.send_keys(Keys.ENTER)
        sleep(TIMER)

        # Entrando nas "Ocorrência"
        entrando_ocorrencias = driver.find_element(By.XPATH, '/html/body/div/div/div/div[1]/div[3]/table/tbody/tr/td[6]/a/span')
        entrando_ocorrencias.click()
        sleep(TIMER)

        # Entrando em "Incluir Ocorrência"
        entrando_incluir = driver.find_element(By.ID, 'addOcorrenciaBtn')
        entrando_incluir.send_keys(Keys.ENTER)
        sleep(TIMER)

        # Seleciona assunto e preenche "13 - Cobrança"
        preenchendo_assunto_1 = driver.find_element(By.ID, 'select2-AssuntoOcorrenciaId-container')
        preenchendo_assunto_1.click()
        preenchendo_assunto_2 = driver.find_element(By.XPATH, '/html/body/span/span/span[1]/input')
        preenchendo_assunto_2.send_keys(cobranca)
        preenchendo_assunto_2.send_keys(Keys.ENTER)

        # Inserindo data
        inserindo_data = driver.find_element(By.ID, 'Data')
        inserindo_data.clear()
        sleep(TIMER)
        inserindo_data.send_keys(str(dia_formatado_preencher))

        # Inserindo "Descrição"
        inserindo_descricao = driver.find_element(By.NAME, 'Descricao')
        inserindo_descricao.send_keys(f'{nome}: {descricao}')

        # Encontra o "Incluir Ocorrência"
        # incluindo = driver.find_element(By.ID, 'createBtn')
        # incluindo.send_keys(Keys.ENTER)
        # sleep(1)
        driver.quit()
        mutuarios_anteriores.append(mutuario)
