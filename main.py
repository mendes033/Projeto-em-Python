# Utilizar pip install selenium
# Utilizar pip install webdriver-manager
# Utilizar pip install openpyxl

import os                  # Trabalhar com pastas windows
from pathlib import Path as p  # Trabalhar com pastas windows
import shutil              # Movimentação de arquivo
import time                # Utilizar no timesleep
import zipfile             # Arquivos compactados
import getpass             # Senha não aparece
import pandas as pd        # Trabalhar com arquivos excel

# Automatização do download de arquivo utilizando Chrome
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

# Solicita crendencial de acesso

userestac = input("Digite o seu usuário: ")
pwordestac = getpass.getpass("Digite a sua senha: ")

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
    time.sleep(3)

    driver.find_element(By.ID, 'titulo-card-entrega-ARA0066').click()
    time.sleep(3)

    driver.find_element(By.XPATH, '//*[@id="temas-lista-topicos"]/li[5]/a/div/div').click()
    time.sleep(3)

    driver.find_element(By.XPATH, '/html/body/div[1]/section/section/main/section/article/div/div[3]/section[1]/section/div[3]/div[2]/div/section/div').click()
    time.sleep(60)

# Fecha o Google Chrome
driver.quit()

# Cria o diretório onde iremos salvar o arquivo baixado.

pasta = p.home() / 'Desktop' / 'ParadigmasTema1'
pasta.mkdir(exist_ok=True)

path = p('Downloads/5m-Sales-Records.zip')

# Pega o arquivo baixado e move para o diretorio que criamos.
# source = '/Users/local/Downloads/5m-Sales-Records.zip'
# destination = Path.home() / 'Desktop/ParadigmasTema1/5m-Sales-Records.zip'
# shutil.move(source, destination)
# time.sleep(5)
#
# # Extrai o arquivo baixado e coloca na mesma pasta que criamos
# extrair = zipfile.ZipFile("C:/Users/201702276368/Vittor-Hugo-Isaac/TrabalhoAltamira-Vittor/combate_pirataria.zip")
# extrair.extractall("C:/Users/201702276368/Vittor-Hugo-Isaac/TrabalhoAltamira-Vittor/")
# extrair.close()
#
# # Leitura do arquivo original
# df = pd.read_csv("C:/Users/201702276368/Vittor-Hugo-Isaac/TrabalhoAltamira-Vittor/Tabela_PACP.csv", sep=";")
#
# # Total de cada coluna com quantidade
# qt_un_apreendidas = df["QT_UN_APREENDIDAS"].sum()
# qt_un_lacradas = df["QT_UN_LACRADAS"].sum()
# qt_un_retidas = df["QT_UN_RETIDAS"].sum()
# qt_un_retiradas = df["QT_UN_RETIRADAS"].sum()
#
# # Quantidade de itens NÃO-NULOS
# qtd_equip_apreendidas = df.QT_UN_APREENDIDAS.count()
# qtd_equip_lacradas = df.QT_UN_LACRADAS.count()
# qtd_equip_retidas = df.QT_UN_RETIDAS.count()
# qtd_equip_retiradas = df.QT_UN_RETIRADAS.count()
#
# # Média aritimética de cada coluna
# media_apreendidas = qt_un_apreendidas / qtd_equip_apreendidas
# media_lacradas = qt_un_lacradas / qtd_equip_lacradas
# media_retidas = qt_un_retidas / qtd_equip_retidas
# media_retiradas = qt_un_retiradas / qtd_equip_retiradas
#
# # Menor valor de cada coluna
# min_apreendidas = df["QT_UN_APREENDIDAS"].min(skipna=True)
# min_lacradas = df["QT_UN_LACRADAS"].min(skipna=True)
# min_retidas = df["QT_UN_RETIDAS"].min(skipna=True)
# min_retiradas = df["QT_UN_RETIRADAS"].min(skipna=True)
#
# # Maior valor de cada coluna
# max_apreendidas = df["QT_UN_APREENDIDAS"].max(skipna=True)
# max_lacradas = df["QT_UN_LACRADAS"].max(skipna=True)
# max_retidas = df["QT_UN_RETIDAS"].max(skipna=True)
# max_retiradas = df["QT_UN_RETIRADAS"].max(skipna=True)
#
# # Cria um novo excel com os dados organizados e os totais
# totais_list = [("TOTAIS:", "", "", qt_un_apreendidas, qt_un_lacradas, qt_un_retidas, "", qt_un_retiradas)]
# qtd_itens = [("TOTAL ITENS:", "", "", qtd_equip_apreendidas, qtd_equip_lacradas, qtd_equip_retidas, "", qtd_equip_retiradas)]
# media_col = [("MEDIA:", "", "", media_apreendidas, media_lacradas, media_retidas, "", media_retiradas)]
# min_val = [("MENOR VALOR:", "", "", min_apreendidas, min_lacradas, min_retidas, "", min_retiradas)]
# max_val = [("MAIOR VALOR:", "", "", max_apreendidas, max_lacradas, max_retidas, "", max_retiradas)]
#
# dfcalc = pd.DataFrame(totais_list,columns=["ID", "Área", "Equipamento", "QT_UN_APREENDIDAS", "QT_UN_LACRADAS", "QT_UN_RETIDAS",
#                                "Valor Estimado", "QT_UN_RETIRADAS"])
# dfqtd = pd.DataFrame(qtd_itens,columns=["ID", "Área", "Equipamento", "QT_UN_APREENDIDAS", "QT_UN_LACRADAS", "QT_UN_RETIDAS",
#                               "Valor Estimado", "QT_UN_RETIRADAS"])
# dfmedia = pd.DataFrame(media_col,columns=["ID", "Área", "Equipamento", "QT_UN_APREENDIDAS", "QT_UN_LACRADAS", "QT_UN_RETIDAS",
#                                 "Valor Estimado", "QT_UN_RETIRADAS"])
# dfmin = pd.DataFrame(min_val,columns=["ID", "Área", "Equipamento", "QT_UN_APREENDIDAS", "QT_UN_LACRADAS", "QT_UN_RETIDAS",
#                               "Valor Estimado", "QT_UN_RETIRADAS"])
# dfmax = pd.DataFrame(max_val,columns=["ID", "Área", "Equipamento", "QT_UN_APREENDIDAS", "QT_UN_LACRADAS", "QT_UN_RETIDAS",
#                               "Valor Estimado", "QT_UN_RETIRADAS"])
#
# df = df.append(dfcalc, ignore_index=True)
# df = df.append(dfqtd, ignore_index=True)
# df = df.append(dfmedia, ignore_index=True)
# df = df.append(dfmin, ignore_index=True)
# df = df.append(dfmax, ignore_index=True)
#
# df.to_excel(r"C:/Users/201702276368/Vittor-Hugo-Isaac/TrabalhoAltamira-Vittor/Resultado_Tabela_PACP.xlsx", index=False,
#             header=True)