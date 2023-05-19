import os
import time
import shutil
import zipfile
import pandas as pd
import smtplib
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv



load_dotenv()
USER_LOGIN = os.getenv('LOGIN')
USER_PASS = os.getenv('PASSWORD')
SITE = os.getenv('URL')

user = "201708220811@alunos.estacio.br"
password_user = "Aa@21051997"


service = Service(executable_path="/usr/local/bin/chromedriver")
with webdriver.Chrome(service=service) as driver:
    driver.get("https://estudante.estacio.br/")
    driver.implicitly_wait(5)

    driver.maximize_window()

    # driver.get("https://estudante.estacio.br/")
    # time.sleep(5)
    # bt0 = driver.find_element_by_xpath("/html/body/div[1]/section/div/div/div/section/div[1]/button").click()
    # time.sleep(15)

    driver.find_element(By.XPATH, "/html/body/div[1]/section/div/div/div/section/div[1]/button").click()
    time.sleep(5)

    # username_textbox=driver.find_element_by_name("loginfmt")
    # username_textbox.send_keys(user)

    username_textbox=driver.find_element(By.NAME, "loginfmt")
    username_textbox.send_keys(user)

    # bt1 = driver.find_element_by_id("idSIButton9").click()
    # time.sleep(5)

    driver.find_element(By.ID , "idSIButton9").click()
    time.sleep(5)

    # password=driver.find_element_by_id("i0118")
    # password.send_keys(password_user)

    password=driver.find_element(By.ID , "i0118")
    password.send_keys(password_user)


    # login_attempt = driver.find_element_by_id("idSIButton9")
    # login_attempt.submit()
    # time.sleep(5)

    login_attempt = driver.find_element(By.ID , "idSIButton9")
    login_attempt.submit()
    time.sleep(5)

    # bt2 = driver.find_element_by_id("idBtn_Back").click()
    # time.sleep(5)

    driver.find_element(By.ID , "idBtn_Back").click()
    time.sleep(5)

    # div/div/div/article/div/div[1]/section/header/button").click()
    driver.find_element(By.XPATH, "/html/body/div[1]/section/section/main/section/article/section[1]/section/div[2]/div/div/section[2]/header/button").click()
    time.sleep(5)


    # bt4 = driver.find_element_by_xpath("/html/body/div[1]/section/article[2]/article[2]/section/div/div/div[1]/section/div/div/article/section[3]/div[2]/div[3]/div/section/div[5]/a").click()
    # time.sleep(5)

    driver.find_element(By.XPATH, "/html/body/div[1]/section/section/main/section/aside/section/section[2]/div[1]/button[2]").click()
    time.sleep(5)

    # bt5 = driver.find_element_by_xpath("/html/body/div[1]/section/article[2]/article[2]/section/div/div/div[2]/section/div/div/article/section/section/section/div[3]/div[2]/div/section/div/div[1]/button").click()
    # time.sleep(5)

    driver.find_element(By.XPATH, "/html/body/div[1]/section/section/main/section/aside/section/section[2]/div[3]/section[2]/section/div[3]/div[2]/div/section/div/div[1]/button").click()
    time.sleep(5)

    # driver.get("https://www.kabum.com.br/")
    # driver.switch_to_window(driver.window_handles[1])
    driver.switch_to.window(driver.window_handles[1])
    driver.get(driver.current_url)
    time.sleep(5)

    # driver.switch_to.new_window("https://www.kabum.com.br/")
    # time.sleep(5)
    # val = "https://www.kabum.com.br/"
    # wait = WebDriverWait(driver, 5)
    # driver.get(val)
    # wait.until(EC.url_to_be(val))
    # page_source = driver.page_source

    # produto = str(input("Digite o nome do produto: "))
    produto = "Fone HyperX"

    busca = driver.find_element(By.ID , "input-busca")
    busca.send_keys(produto)

    busca_attempt = driver.find_element(By.XPATH, "/html/body/div[1]/header/div[1]/div/div/div[1]/div[3]/div/form/button").click()
    time.sleep(5)

# driver.quit()


# os.makedirs("./TrabalhoAltamira", exist_ok=True)


# source = r"C:/Users/201801196982/Downloads/combate_pirataria.zip"
# destination = r"C:/Users/201801196982/PycharmProjects/pythonProject/combate_pirataria.zip"
# shutil.move(source, destination)
# time.sleep(5)

# #Extrai o arquivo baixado e coloca na mesma pasta que criamos
# extrair = zipfile.ZipFile("C:/Users/201801196982/PycharmProjects/pythonProject/combate_pirataria.zip")
# extrair.extractall("C:/Users/201801196982/PycharmProjects/pythonProject")
# extrair.close()

# #Leitura do arquivo original
# df = pd.read_csv("C:/Users/201801196982/PycharmProjects/pythonProject/Tabela_PACP.csv", sep=";")










# #Total de cada coluna com quantidade
# qt_un_apreendidas = df["QT_UN_APREENDIDAS"].sum()
# qt_un_lacradas = df["QT_UN_LACRADAS"].sum()
# qt_un_retidas = df["QT_UN_RETIDAS"].sum()
# qt_un_retiradas = df["QT_UN_RETIRADAS"].sum()

# #Quantidade de itens NÃO-NULOS
# qtd_equip_apreendidas = df.QT_UN_APREENDIDAS.count()
# qtd_equip_lacradas = df.QT_UN_LACRADAS.count()
# qtd_equip_retidas = df.QT_UN_RETIDAS.count()
# qtd_equip_retiradas = df.QT_UN_RETIRADAS.count()

# #Média aritimética de cada coluna
# media_apreendidas = qt_un_apreendidas/qtd_equip_apreendidas
# media_lacradas = qt_un_lacradas/qtd_equip_lacradas
# media_retidas = qt_un_retidas/qtd_equip_retidas
# media_retiradas = qt_un_retiradas/qtd_equip_retiradas

# #Menor valor de cada coluna
# min_apreendidas = df["QT_UN_APREENDIDAS"].min(skipna=True)
# min_lacradas = df["QT_UN_LACRADAS"].min(skipna=True)
# min_retidas = df["QT_UN_RETIDAS"].min(skipna=True)
# min_retiradas = df["QT_UN_RETIRADAS"].min(skipna=True)

# #Maior valor de cada coluna
# max_apreendidas = df["QT_UN_APREENDIDAS"].max(skipna=True)
# max_lacradas = df["QT_UN_LACRADAS"].max(skipna=True)
# max_retidas = df["QT_UN_RETIDAS"].max(skipna=True)
# max_retiradas = df["QT_UN_RETIRADAS"].max(skipna=True)

# #Cria um novo excel com os dados organizados e os totais
# totais_list = [("TOTAIS:","","",qt_un_apreendidas,qt_un_lacradas,qt_un_retidas,"",qt_un_retiradas)]
# qtd_itens = [("TOTAL ITENS:","","",qtd_equip_apreendidas,qtd_equip_lacradas,qtd_equip_retidas,"",qtd_equip_retiradas)]
# media_col = [("MEDIA:","","",media_apreendidas,media_lacradas,media_retidas,"",media_retiradas)]
# min_val = [("MENOR VALOR:","","",min_apreendidas,min_lacradas,min_retidas,"",min_retiradas)]
# max_val = [("MAIOR VALOR:","","",max_apreendidas,max_lacradas,max_retidas,"",max_retiradas)]

# dfcalc = pd.DataFrame(totais_list, columns=["ID", "Área","Equipamento","QT_UN_APREENDIDAS","QT_UN_LACRADAS","QT_UN_RETIDAS","Valor Estimado","QT_UN_RETIRADAS"])
# dfqtd = pd.DataFrame(qtd_itens, columns=["ID", "Área","Equipamento","QT_UN_APREENDIDAS","QT_UN_LACRADAS","QT_UN_RETIDAS","Valor Estimado","QT_UN_RETIRADAS"])
# dfmedia = pd.DataFrame(media_col, columns=["ID", "Área","Equipamento","QT_UN_APREENDIDAS","QT_UN_LACRADAS","QT_UN_RETIDAS","Valor Estimado","QT_UN_RETIRADAS"])
# dfmin = pd.DataFrame(min_val, columns=["ID", "Área","Equipamento","QT_UN_APREENDIDAS","QT_UN_LACRADAS","QT_UN_RETIDAS","Valor Estimado","QT_UN_RETIRADAS"])
# dfmax = pd.DataFrame(max_val, columns=["ID", "Área","Equipamento","QT_UN_APREENDIDAS","QT_UN_LACRADAS","QT_UN_RETIDAS","Valor Estimado","QT_UN_RETIRADAS"])

# df = df.append(dfcalc,ignore_index=True)
# df = df.append(dfqtd,ignore_index=True)
# df = df.append(dfmedia,ignore_index=True)
# df = df.append(dfmin,ignore_index=True)
# df = df.append(dfmax,ignore_index=True)

# df.to_excel(r"C:/Users/201801196982/PycharmProjects/pythonProject/ResTabela.xlsx", index=False, header=True)


# host = "smtp.office365.com"
# port = 587
# user_email = input("Informe o e-mail de acesso:")
# password_email = input("Informe a senha de acesso:")

# server = smtplib.SMTP(host, port)

# server.ehlo()
# server.starttls()
# server.login(user_email, password_email)


# email_msg = MIMEMultipart()
# email_msg["From"] = user_email
# email_msg["to"] = ""
# email_msg["Subject"] = "Trabalho - Paradigmas de Linguagens de Programação em Python"

# menssage = """<p>Segue em anexo</p>
# <p>Atenciosamente,</p>
# <p>Fabio Eduardo</p> &
# <p>João Vitor </p>
# <p>Matricula: 201801196982</p>
# <p>Matricula: </p>
# """
# email_msg.attach(MIMEText(menssage, "html"))

# anexo_arq = "C:/Users/201801196982/PycharmProjects/pythonProject/ResTabela.xlsx"
# attachment = open(anexo_arq, "rb")

# att = MIMEBase("application","octet-stream")
# att.set_payload(attachment.read())
# encoders.encode_base64(att)

# att.add_header("Content-Disposition", "attachment; filename=Resultado_Tabela.xlsx")
# attachment.close()
# email_msg.attach(att)

# #Efetua o envio do email
# server.sendmail(email_msg["From"], email_msg["To"], email_msg.as_string())
# print("Email enviado com sucesso!")

#Finaliza a autenticação com o servidor.
# server.quit()
