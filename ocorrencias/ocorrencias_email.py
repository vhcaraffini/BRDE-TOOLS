from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from tkinter import filedialog
from selenium import webdriver
import tkinter as tk
import pandas as pd

TIMER = 15


def registrar_ocorrencia_email(data):
    # Dia do registro
    dia_formatado_preencher = data

    # Abrindo Excel
    root = tk.Tk()
    root.withdraw() # Esconde a janela principal

    CAMINHO_ARQUIVO = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    df = pd.read_excel(CAMINHO_ARQUIVO, sheet_name='Planilha1')

    # Valores
    mutuarios_anteriores = []

    # Pegando valores do excel e adicionando a lista
    for i, mutuario in enumerate(df['MUTUÁRIO']):
        nome = 'Cris'
        cobranca = '13'
        descricao = df.loc[i, 'OBSERVAÇÃO']

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

        # Entrando nas "Ocorrência"
        entrando_ocorrencias = WebDriverWait(driver, TIMER).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div/div[1]/div[3]/table/tbody/tr/td[6]/a/span')))
        entrando_ocorrencias.click()

        # Entrando em "Incluir Ocorrência"
        entrando_incluir = WebDriverWait(driver, TIMER).until(EC.presence_of_element_located((By.ID, 'addOcorrenciaBtn')))
        entrando_incluir.send_keys(Keys.ENTER)

        # Seleciona assunto e preenche "13 - Cobrança"
        preenchendo_assunto_1 = WebDriverWait(driver, TIMER).until(EC.presence_of_element_located((By.ID, 'select2-AssuntoOcorrenciaId-container')))
        preenchendo_assunto_1.click()
        preenchendo_assunto_2 = WebDriverWait(driver, TIMER).until(EC.presence_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input')))
        preenchendo_assunto_2.send_keys(cobranca)
        preenchendo_assunto_2.send_keys(Keys.ENTER)

        # Inserindo data
        inserindo_data = WebDriverWait(driver, TIMER).until(EC.presence_of_element_located((By.ID, 'Data')))
        inserindo_data.clear()
        inserindo_data.send_keys(str(dia_formatado_preencher))

        # Inserindo "Descrição"
        inserindo_descricao = WebDriverWait(driver, TIMER).until(EC.presence_of_element_located((By.NAME, 'Descricao')))
        inserindo_descricao.send_keys(f'{nome}: {descricao}')

        # Encontra o "Incluir Ocorrência"
        incluindo = WebDriverWait(driver, TIMER).until(EC.presence_of_element_located((By.ID, 'createBtn')))
        incluindo.send_keys(Keys.ENTER)
        driver.quit()
        mutuarios_anteriores.append(mutuario)