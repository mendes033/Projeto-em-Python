import shutil                              # Movimentação de arquivo
import time                                # Utilizar no timesleep
import getpass                             # Senha não aparece
import pandas as pd                        # Trabalhar com arquivos excel
from py7zr import SevenZipFile as szf      # Arquivos compactados
from pathlib import Path as p              # Trabalhar com pastas windows, independente do usuário ou so utilizado

# Automatização do download de arquivo utilizando Chrome

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

# Solicita crendencial de acesso

userestac = input("Digite o seu usuário: ")
pwordestac = getpass.getpass("Digite a sua senha: ")

# Inicia o google chrome, acessa a materia, conteudo complementar, clica no download e fecha o chrome

service = Service(executable_path="/usr/local/bin/chromedriver")
with webdriver.Chrome(service=service) as driver:
    driver.get("https://estudante.estacio.br/")
    driver.implicitly_wait(2)
    driver.maximize_window()

    driver.find_element(By.XPATH, '//*[@id="section-login"]/div/div/div[1]/section/div[1]/button').click()
    time.sleep(2)

    username_textbox=driver.find_element(By.XPATH, '//*[@id="i0116"]')
    username_textbox.send_keys(userestac)

    driver.find_element(By.XPATH, '//*[@id="idSIButton9"]').click()
    time.sleep(2)

    password=driver.find_element(By.XPATH, '//*[@id="i0118"]')
    password.send_keys(pwordestac)

    login_attempt = driver.find_element(By.XPATH, '//*[@id="idSIButton9"]')
    login_attempt.submit()
    time.sleep(2)

    driver.find_element(By.XPATH, '//*[@id="idBtn_Back"]').click()
    time.sleep(5)

    driver.find_element(By.ID, 'titulo-card-entrega-ARA0066').click()
    time.sleep(3)

    driver.find_element(By.XPATH, '//*[@id="temas-lista-topicos"]/li[5]/a/div/div').click()
    time.sleep(3)

    driver.find_element(By.XPATH, '/html/body/div[1]/section/section/main/section/article/div/div[3]/section[1]/section/div[3]/div[2]/div/section/div').click()
    time.sleep(60)

# Fecha o Google Chrome
driver.quit()

# Cria o pasta onde iremos salvar o arquivo baixado e verifica se o arquivo e a pasta já existem.

pastadesk = p.home() / 'Desktop' / 'ParadigmasTema1'
pastadesk.mkdir(exist_ok=True)

# Pega o arquivo baixado e move para pasta que foi criada.

shutil.move(p.home() /'Downloads'/'5m-Sales-Records.7z', p.home() /'Desktop'/'ParadigmasTema1'/'5m-Sales-Records.7z')

# Extração do 7z

with szf(p.home() /'Desktop'/'ParadigmasTema1'/'5m-Sales-Records.7z', mode='r') as z:
    z.extractall(p.home() /'Desktop'/'ParadigmasTema1')

# Loop da leitura do csv e exportação a cada 500.000 linhas, separando as planilhas automaticamente.

nrows = 1

i=1

for nrows in range (5000000):

    df = pd.read_csv(p.home() / 'Desktop' / 'ParadigmasTema1' / '5m Sales Records.csv',
                     names=['Region', 'Country', 'Item_Type', 'Sales_Channel', 'Order_Priority', 'Order_Date','Order_ID', 'Ship_Date',
                            'Units_Sold', 'Unit_Price', 'Unit_Cost', 'Total_Revenue', 'Total_Cost', 'Total_Profit'],
                     skiprows=nrows, nrows=500000)

    df.fillna('0', inplace=True)

    df = df.sort_values(['Country'])

    df.to_excel(p.home() / 'Desktop' / 'ParadigmasTema1' / ('tema1-pronto'+str(i)+'.xlsx'), index=False)

    i+=1

    if nrows==1:
        nrows=500000

    else:
        nrows += 500000