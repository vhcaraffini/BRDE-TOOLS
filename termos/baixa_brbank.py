from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
from selenium.webdriver.support.ui import Select
import tkinter as tk
from tkinter import filedialog
import pandas as pd

TIMER = 0.5

def alterando_situacao_contrato():
    # Abrindo Excel
    root = tk.Tk()
    root.withdraw()

    CAMINHO_ARQUIVO = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    df = pd.read_excel(CAMINHO_ARQUIVO, sheet_name='Planilha1')

    for i, mutuario in enumerate(df['Número do contrato']):
        ...
    # # Listas
    # mutuarios = []
    # valores = []
    # datas = []

    # Pegando valores do excel e adicionando a lista
    # for row in sheet.iter_rows(values_only=True):
    #     if row[0] == 'EMPRESA':
    #         continue

    #     elif row[0] == 'TOTAL - TARIFA DE CADASTRO':
    #         break

    #     mutuarios.append(row[0])
    #     valores.append(str(row[1]) + '00')
    #     datas.append(row[2])

    # mutuarios
    for n in range(1):
        # mutuario = mutuarios[n].strip()
        # valor = valores[n]
        # data = datas[n].strftime('%d/%m/%Y')

        # Acessando o site
        driver = webdriver.Chrome()
        driver.get(f"https://brbank.brde.com.br/Contratos/Buscar")

        # Entrando no Quitar contratos
        quitar_contratos = driver.find_element(By.ID, 'quitarContratosBtn')
        quitar_contratos.click()
        sleep(TIMER)

        # Pesquisando pelo nome
        pesquisando_por_nome = driver.find_element(By.ID, 'nomeCnpjCpfModal')
        pesquisando_por_nome.send_keys('Enio Jose Seganfredo')
        sleep(TIMER)
        pesquisando = driver.find_element(By.XPATH, '//*[@id="searchPessoaBtn"]/span')
        pesquisando.click()
        sleep(TIMER)

        # Selecionando bolinha
        bolinha_de_seleção = driver.find_element(By.NAME, 'radioPessoa')
        bolinha_de_seleção.click()

        # Dando Ok
        dand_ok = driver.find_element(By.ID, 'addBtn')
        dand_ok.click()
        sleep(TIMER)

        # Entrando no lapiz
        lapiz_edicao = driver.find_element(By.XPATH, '//*[@id="quitacaoContratosTable"]/tbody/tr/td[9]/div/a[1]/span')
        lapiz_edicao.click()
        sleep(TIMER)

        # Modificando situação
        modifica_situacao = driver.find_element(By.XPATH, "/html/body/div/div/div/div/form/fieldset/section[2]/div/table/tbody/tr/td[8]/span/span[1]/span")
        modifica_situacao.click()
        modifica_situacao.send_keys(Keys.ARROW_DOWN)
        modifica_situacao.send_keys(Keys.ENTER)
        sleep(TIMER)

        # Preenche numero do termo
        preenche_numero_termo = driver.find_element(By.XPATH, '/html/body/div/div/div/div/form/fieldset/section[2]/div/table/tbody/tr/td[6]/input')
        preenche_numero_termo.send_keys('12323')

        # Preenche data do termo
        preenche_data_termo = driver.find_element(By.XPATH, '/html/body/div/div/div/div/form/fieldset/section[2]/div/table/tbody/tr/td[6]/input')
        preenche_data_termo.send_keys('12/03/2023')

        # Encontra o Salvar
        ending = driver.find_element(By.XPATH, '/html/body/div/div/div/div/form/fieldset/section[2]/div/table/tbody/tr/td[9]/div/a[2]/span')
        # ending.click()

        # # Salva
        # atualizar = driver.find_element(By.XPATH, 'select2-SituacaoId-container') 
        # atualizar.send_keys(Keys.ENTER)
        # sleep(3)

        driver.quit()