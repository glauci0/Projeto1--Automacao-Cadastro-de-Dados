# Importar as Bibliotecas necessárias
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook
from datetime import date
import time
# ------------------------------------------

openArquivo = load_workbook(r'C:\Users\glauc\OneDrive\Documentos\Robô_SAS\Automacao_SAS\Validacaoforms.xlsx') #Leitura do arquivo para o openpyxl
act = openArquivo.active #Apontamento de planilha para o openpyxl   

arquivo = pd.read_excel(r'C:\Users\glauc\OneDrive\Documentos\Robô_SAS\Automacao_SAS\Validacaoforms.xlsx') #Leitura do arquivo para o pandas
print(arquivo) #Leitura da planilha no estado inicial
link = "https://docs.google.com/forms/d/e/1FAIpQLSeVfnHz95o4A3XEJWzlmvpD-oasKZzUt-p1bRf9368t0gX5vQ/viewform?usp=header" #Link Forms

#Configuração para o chrome abrir em segundo plano
opc = webdriver.ChromeOptions()
# opc.add_argument("--headless") tirar o "#" para atuar em segundo plano
# ------------------------------------------

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service = service, options=opc)

driver.implicitly_wait(15) #Declara uma espera implícita, caso ação/processo não for executada em 15 segundos ele aborta o processo
driver.get(link) #Abrir o link definido

time.sleep(2)

for i, row in arquivo.iterrows(): #Loop para percorrer toda linha do arquivo e fazer determinadas ações
    print("Executando...")
    #Variávei de uso
    nome = row['nome'] #Variável de nome
    nascimento = str(row['nascimento']).split(' ')[0].split('-') # Transforma a data em data arrey na variável
    dia = nascimento[2] #Pega o dia do data arrey para a variável
    mes = nascimento[1] #Pega o mês do data arrey para a variável
    ano = nascimento[0] #Pega o ano do data arrey para a variável
    ano_atual = date.today().year
    idade = ano_atual - int(ano) # faz o calculo da idade, utilizando o ano de nascimento menos ano atual
    dataNascimento = dia + mes + ano
    cargo = row['cargo'] # Variável cargo funcionário
    linha = i + 2 # Variável para localização de linha openpyxl
    # ------------------------------------------
    # Envia variáveis para inputs definidos
    driver.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(nome)
    time.sleep(1)
    driver.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div/div[2]/div[1]/div/div[1]/input').send_keys(dataNascimento)
    time.sleep(1)
    driver.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(idade)
    time.sleep(1)
    driver.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(cargo)
    # ------------------------------------------
    act[f'E{linha}'] = "OK" # Escreve OK na planilha com openpyxl para saber os nomes escritos
    act[f'D{linha}'] = driver.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[1]/div/div[4]/div[3]').get_attribute('innerText') # 'raspagem' escreve o atributo da web

    driver.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span/span').click() # Clicar em enviar o forms
    time.sleep(1)
    driver.find_element(By.XPATH,'/html/body/div[1]/div[2]/div[1]/div/div[4]/a').click() # Clicar para enviar outra resposta
    openArquivo.save('Validacaoforms.xlsx') # Salva planilha

driver.quit() # fecha o driver
arquivo = pd.read_excel('Validacaoforms.xlsx') # Leitura do arquivo para o pandas
print(arquivo) # Print da planilha no estado final