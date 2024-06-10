from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta 
import openpyxl

TIMER = 10

# Abrindo Excel
workbook = openpyxl.load_workbook('C:/Users/e.marcus.machado/OneDrive - Banco Regional de Desenvolvimento do Extremo Sul/Pasta_teste.xlsx')
sheet = workbook['Planilha1']

# Listas
mutuarios = []
valores = []
datas = []

# Pegando valores do excel e adicionando a lista
for row in sheet.iter_rows(values_only=True):
    if row[0] == 'EMPRESA':
        continue

    elif row[0] == 'TOTAL - TARIFA DE CADASTRO':
        break

    mutuarios.append(row[0])
    valores.append(str(row[1]) + '00')
    datas.append(row[2])

for n in range(len(mutuarios)):
    mutuario = mutuarios[n].strip()
    valor = valores[n]
    data = datas[n].strftime('%d/%m/%Y')

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