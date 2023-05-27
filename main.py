# Utilizar pip install selenium
# Utilizar pip install webdriver-manager
# Utilizar pip install openpyxl
# Utilizar pip install py7zr

from pathlib import Path as p  # Trabalhar com pastas windows, independente do usuário ou so utilizado
import shutil                  # Movimentação de arquivo
import time                    # Utilizar no timesleep
import py7zr                # Arquivos compactados
import getpass                 # Senha não aparece
import pandas as pd            # Trabalhar com arquivos excel

# Automatização do download de arquivo utilizando Chrome
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

# Solicita crendencial de acesso

# userestac = input("Digite o seu usuário: ")
# pwordestac = getpass.getpass("Digite a sua senha: ")

# # Inicia o google chrome, acessa a materia, conteudo complementar, clica no download e fecha o chrome

# service = Service(executable_path="/usr/local/bin/chromedriver")
# with webdriver.Chrome(service=service) as driver:
#     driver.get("https://estudante.estacio.br/")
#     driver.implicitly_wait(2)
#     driver.maximize_window()

#     driver.find_element(By.XPATH, '//*[@id="section-login"]/div/div/div[1]/section/div[1]/button').click()
#     time.sleep(2)

#     username_textbox=driver.find_element(By.XPATH, '//*[@id="i0116"]')
#     username_textbox.send_keys(userestac)

#     driver.find_element(By.XPATH, '//*[@id="idSIButton9"]').click()
#     time.sleep(2)

#     password=driver.find_element(By.XPATH, '//*[@id="i0118"]')
#     password.send_keys(pwordestac)

#     login_attempt = driver.find_element(By.XPATH, '//*[@id="idSIButton9"]')
#     login_attempt.submit()
#     time.sleep(2)

#     driver.find_element(By.XPATH, '//*[@id="idBtn_Back"]').click()
#     time.sleep(3)

#     driver.find_element(By.ID, 'titulo-card-entrega-ARA0066').click()
#     time.sleep(3)

#     driver.find_element(By.XPATH, '//*[@id="temas-lista-topicos"]/li[5]/a/div/div').click()
#     time.sleep(3)

#     driver.find_element(By.XPATH, '/html/body/div[1]/section/section/main/section/article/div/div[3]/section[1]/section/div[3]/div[2]/div/section/div').click()
#     time.sleep(60)

# # Fecha o Google Chrome
# driver.quit()

# # Cria o pasta onde iremos salvar o arquivo baixado e verifica se o arquivo e a pasta já existem.

# pastadesk = p.home() / 'OneDrive' / 'Desktop' / 'ParadigmasTema1'
# pastadesk.mkdir(exist_ok=True)

# Pega o arquivo baixado e move para pasta que foi criada.

# shutil.move(p.home() /'Downloads'/'5m-Sales-Records.7z', p.home() / 'OneDrive' /'Área de Trabalho'/'ParadigmasTema1'/'5m-Sales-Records.7z')

# Leitura do arquivo original
# pd.read_csv(p.home() /'Desktop'/'ParadigmasTema1'/'5m-Sales-Records.7z'/ '5m-Sales-Records.csv', sep=",")

with py7zr.SevenZipFile(p.home() / 'OneDrive' /'Área de Trabalho'/'ParadigmasTema1'/'5m-Sales-Records.7z') as z:
    with z.open('5m-Sales-Records.csv') as f:
        df = pd.read_csv(f, columns=['Region', 'Country', 'Item_Type', 'Sales_Channel', 'Order_Priority', 'Order_Date', 'Order_ID','Ship_Date', 'Units_Sold', 'Unit_Price', 'Unit_Cost', 'Total_Revenue', 'Total_Cost', 'Total_Profit'])

# Cria o dataframe onde vai armazenar os dados que foram lidos

df.fillna('0', inplace=True)
print(df)

# Cria um novo excel com os dados organizados