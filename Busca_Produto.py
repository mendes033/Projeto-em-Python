import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from dotenv import load_dotenv
from pathlib import Path as p
import getpass


class Produto:
    def __init__(self, nome, valor_real, valor_promo, temp_promo, link):
        self.nome = nome
        self.valor_real = valor_real
        self.valor_promo = valor_promo
        self.temp_promo = temp_promo
        self.link = link

load_dotenv()
USER_LOGIN = os.getenv('LOGIN')
USER_PASS = os.getenv('PASSWORD')
SITE = os.getenv('URL')

user = input("Digite o seu usuário: ")
password_user = getpass.getpass("Digite a sua senha: ")

filter = str(input("Digite o nome do produto: "))


service = Service(executable_path="/usr/local/bin/chromedriver")
with webdriver.Chrome(service=service) as driver:
    driver.get("https://estudante.estacio.br/")
    driver.implicitly_wait(5)
    driver.maximize_window()

    driver.find_element(By.XPATH, "/html/body/div[1]/section/div/div/div/section/div[1]/button").click()
    time.sleep(5)

    username_textbox=driver.find_element(By.NAME, "loginfmt")
    username_textbox.send_keys(user)

    driver.find_element(By.ID , "idSIButton9").click()
    time.sleep(5)

    password=driver.find_element(By.ID , "i0118")
    password.send_keys(password_user)

    login_attempt = driver.find_element(By.ID , "idSIButton9")
    login_attempt.submit()
    time.sleep(5)

    driver.find_element(By.ID , "idBtn_Back").click()
    time.sleep(5)

    driver.find_element(By.ID, 'titulo-card-entrega-ARA0066').click()
    time.sleep(5)

    driver.find_element(By.XPATH, '//*[@id="temas-lista-topicos"]/li[5]/a/div/div').click()
    time.sleep(5)

    driver.find_element(By.XPATH, "/html/body/div[1]/section/section/main/section/article/div/div[3]/section[2]/section/div[3]/div[2]/div/section/div/div[1]/button").click()
    time.sleep(5)

    driver.switch_to.window(driver.window_handles[1])
    driver.get(driver.current_url)
    time.sleep(5)

    busca = driver.find_element(By.ID , "input-busca")
    busca.send_keys(filter)

    busca_attempt = driver.find_element(By.XPATH, "/html/body/div[1]/header/div[1]/div/div/div[1]/div[3]/div/form/button").click()
    time.sleep(5)

    lista_produtos = []
    i = 0

    while i>=0:
        i = i + 1
        path_nome = "/html/body/div[1]/div[1]/div[3]/div/div/div[2]/div/main/div["+str(i)+"]/a/div/button/div/h2/span"
        path_valor_real = "/html/body/div[1]/div[1]/div[3]/div/div/div[2]/div/main/div["+str(i)+"]/a/div/div[1]/span[1]"
        path_valor_promo = "/html/body/div[1]/div[1]/div[3]/div/div/div[2]/div/main/div["+str(i)+"]/a/div/div[1]/span[2]"
        path_temp_promo = "/html/body/div[1]/div[1]/div[3]/div/div/div[2]/div/main/div["+str(i)+"]/div[2]/div[2]/div/div[2]/div/span[contains(@class,'countdownOffer')]"
        path_link = "/html/body/div[1]/div[1]/div[3]/div/div/div[2]/div/main/div["+str(i)+"]/a"

        nome_produto = driver.find_elements(By.XPATH, path_nome)
        valor_real = driver.find_elements(By.XPATH, path_valor_real)
        valor_promo = driver.find_elements(By.XPATH, path_valor_promo)
        temp_promo = driver.find_elements(By.XPATH, path_temp_promo)
        links = driver.find_elements(By.XPATH, path_link)
        link = ''
        tp = ''

        for l in links:
            link = l.get_attribute('href')

        
        for tp in temp_promo:
            tp = tp.text

        try:
            lista_produtos.append(
                Produto(nome_produto[0].text,
                valor_real[0].text.replace('R$ ', '').replace('.', '').replace(',', '.') or valor_promo[0].text.replace('R$ ', '').replace('.', '').replace(',', '.'),
                valor_promo[0].text.replace('.', '').replace(',', '.') if valor_real[0].text.replace('R$ ', '').replace('.', '').replace(',', '.') else "",
                tp,
                link
                ))
        except:
            i = -1

df = pd.DataFrame(columns=['Nome', 'Valor Real', 'Valor Promo', 'Duração Promo', 'Link'])

for produto in lista_produtos:
   df.loc[len(df)] = [produto.nome, produto.valor_real, produto.valor_promo, produto.temp_promo, produto.link]

df['Valor Real'] = pd.to_numeric(df['Valor Real'])
df['Valor Real'] = df['Valor Real'].astype(float)

df = df.sort_values('Valor Real')
df['Valor Real'] = df['Valor Real'].astype(str)
df['Valor Promo'] = df['Valor Promo'].astype(str)

df['Valor Real'] = 'R$ ' + df['Valor Real'].map(str)

print(df)

folder_path = p.home() / 'Imagens' / (filter + '.xlsx')
df.to_excel(folder_path, index=False, header=True)