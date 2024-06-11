from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from datetime import datetime, timedelta 
from tkinter import filedialog
from selenium import webdriver
from time import sleep
import tkinter as tk
import pandas as pd


TIMER = 10

def tarifa_de_cadastro():

    root = tk.Tk()
    root.withdraw() # Esconde a janela principal

    CAMINHO_ARQUIVO = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    df = pd.read_excel(CAMINHO_ARQUIVO, sheet_name='Planilha1')

    for i, mutuario in enumerate(df['MUTUÁRIO']):
        data = df.loc[i, '']
        valor = df.loc[i, '']
        # Mostrando Mutuario
        print(mutuario)

        # Acessando o site
        driver = webdriver.Chrome()
        driver.get(f"https://brbank.brde.com.br/Pessoas/Buscar")

        # Encontrando e preenchendo barra de pesquisa
        search_input = WebDriverWait(driver, TIMER).until(EC.presence_of_element_located((By.ID, 'NomeCnpjCpf')))
        search_input.send_keys(mutuario)
        search_input.send_keys(Keys.ENTER)

        # Entrando no editar
        entrando_lapis = WebDriverWait(driver, TIMER).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div/div[1]/div[3]/table/tbody/tr/td[4]/a/span')))
        entrando_lapis.click()

        # Apaga e insere valor
        apaga_insere_valor = WebDriverWait(driver, TIMER).until(EC.presence_of_element_located((By.NAME, 'ValorTaxaCadastro')))
        apaga_insere_valor.clear()
        apaga_insere_valor.send_keys(valor)

        # Apaga e insere data
        insere_data = WebDriverWait(driver, TIMER).until(EC.presence_of_element_located((By.NAME, 'DataPagamentoTaxaCadastro')))
        insere_data.clear()
        insere_data.send_keys(data)

        # Encontra o Salvar
        ending = WebDriverWait(driver, TIMER).until(EC.presence_of_element_located((By.ID, 'createBtn')))
        ending.send_keys(Keys.ENTER)

        # Salva
        atualizar = WebDriverWait(driver, TIMER).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/div[2]/button[1]'))) 
        atualizar.send_keys(Keys.ENTER)

        sleep(1.5)
        driver.quit()