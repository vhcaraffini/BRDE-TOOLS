from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
from tkinter import filedialog
from time import sleep
import tkinter as tk
import pandas as pd

TIMER = 0.75

def alterando_situacao_contrato():
    # Abrindo Excel
    root = tk.Tk()
    root.withdraw()

    CAMINHO_ARQUIVO = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo", filetypes=(("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    df = pd.read_excel(CAMINHO_ARQUIVO, sheet_name='Planilha1')

    for i, contrato in enumerate(df['CONTRATO']):
        cgc = str(df.loc[i, 'CPF/CNPJ'])
        numero_termo = str(df.loc[i, 'NR_TERMO_QUITACAO'])
        data_termo = df.loc[i, 'DT_QUITACAO']

        # Acessando o site
        driver = webdriver.Chrome()
        driver.get(f"https://brbank.brde.com.br/Contratos/Buscar")

        # Entrando no Quitar contratos
        quitar_contratos = driver.find_element(By.ID, 'quitarContratosBtn')
        quitar_contratos.click()
        sleep(TIMER)

        # Pesquisando pelo nome
        pesquisando_por_cgc = driver.find_element(By.ID, 'nomeCnpjCpfModal')
        pesquisando_por_cgc.send_keys(str(cgc))
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

        n = 0
        m = 4
        contrato_site = ''
        while str(contrato_site) != str(contrato):
            n += 1
            if n == 10:
                proxima_pagina_garantia = driver.find_element(By.XPATH, F'/html/body/div/div/div/div/form/fieldset/section[2]/nav/ul/li[{m}]/a')
                proxima_pagina_garantia.click()
                sleep(TIMER)
                n = 0
                if m < 6:
                    m += 1
            try:
                encontrando_contrato = driver.find_element(By.XPATH, f'/html/body/div/div/div/div/form/fieldset/section[2]/div/table/tbody/tr[{n}]/td[1]')
                contrato_site = encontrando_contrato.text
            except:
                encontrando_contrato = driver.find_element(By.XPATH, f'/html/body/div/div/div/div/form/fieldset/section[2]/div/table/tbody/tr/td[1]')
                contrato_site = encontrando_contrato.text

            print(f'site:{contrato_site}')
            print(f'Planilha:{contrato}')
            print(n)
        # Entrando no lapiz
        try:
            lapiz_edicao = driver.find_element(By.XPATH, f'/html/body/div/div/div/div/form/fieldset/section[2]/div/table/tbody/tr[{n}]/td[9]/div/a[1]/span')
            lapiz_edicao.click()
        except:
            lapiz_edicao = driver.find_element(By.XPATH, '/html/body/div/div/div/div/form/fieldset/section[2]/div/table/tbody/tr/td[9]/div/a[1]/span')
            lapiz_edicao.click()
        sleep(TIMER)

        # Modificando situação
        try:
            modifica_situacao = driver.find_element(By.XPATH, f"/html/body/div/div/div/div/form/fieldset/section[2]/div/table/tbody/tr[{n}]/td[8]/span/span[1]/span")
            modifica_situacao.click()
            modifica_situacao.send_keys(Keys.ARROW_DOWN)
            modifica_situacao.send_keys(Keys.ENTER)
        except:
            modifica_situacao = driver.find_element(By.XPATH, "/html/body/div/div/div/div/form/fieldset/section[2]/div/table/tbody/tr/td[8]/span/span[1]/span")
            modifica_situacao.click()
            modifica_situacao.send_keys(Keys.ARROW_DOWN)
            modifica_situacao.send_keys(Keys.ENTER)
        sleep(TIMER)

        # Preenche numero do termo
        try:
            preenche_numero_termo = driver.find_element(By.XPATH, f'/html/body/div/div/div/div/form/fieldset/section[2]/div/table/tbody/tr[{n}]/td[6]/input')
            preenche_numero_termo.send_keys(numero_termo)
        except:
            preenche_numero_termo = driver.find_element(By.XPATH, '/html/body/div/div/div/div/form/fieldset/section[2]/div/table/tbody/tr/td[6]/input')
            preenche_numero_termo.send_keys(numero_termo)
        sleep(TIMER)
        
        # Preenche data do termo
        try:
            preenche_data_termo = driver.find_element(By.XPATH, f'/html/body/div/div/div/div/form/fieldset/section[2]/div/table/tbody/tr[{n}]/td[7]/input')
            preenche_data_termo.send_keys(data_termo)
        except:
            preenche_data_termo = driver.find_element(By.XPATH, '/html/body/div/div/div/div/form/fieldset/section[2]/div/table/tbody/tr/td[7]/input')
            preenche_data_termo.send_keys(data_termo)
        sleep(TIMER)

        # Encontra o Salvar
        try:
            ending = driver.find_element(By.XPATH, f'/html/body/div/div/div/div/form/fieldset/section[2]/div/table/tbody/tr[{n}]/td[9]/div/a[2]/span')
            ending.click()
        except:
            ending = driver.find_element(By.XPATH, '/html/body/div/div/div/div/form/fieldset/section[2]/div/table/tbody/tr/td[9]/div/a[2]/span')
            ending.click()
        sleep(TIMER)

        driver.quit()