import pandas as pd
import json
import requests
from datetime import datetime
import time
import smtplib
import datetime as dt





from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#https://economia.awesomeapi.com.br/last/USD-BRL,EUR-BRL,BTC-BRL

while True:
    cotacoes = requests.get("https://economia.awesomeapi.com.br/last/USD-BRL,EUR-BRL,BTC-BRL")
    cotacoes = cotacoes.json()
    dolar = cotacoes["USDBRL"]["bid"]
    euro = cotacoes["EURBRL"]["bid"]
    bitcoin = cotacoes["BTCBRL"]["bid"]


    
    
    #tratamento dos dados e o preenchimento da planilha


    tabela = pd.read_excel("Cotacao.xlsx")
    tabela.loc[0,"Cotação"] = float(dolar)
    tabela.loc[1,"Cotação"] = float(euro)
    tabela.loc[2,"Cotação"] = float(bitcoin)
    tabela.loc[0,"Data Última Atualização"] = datetime.now()
    tabela.loc[1,"Data Última Atualização"] = datetime.now()
    tabela.loc[2,"Data Última Atualização"] = datetime.now()

    tabela.to_excel("Cotacao.xlsx", index=False)
    print(f"Cotação Atualizada. {datetime.now()}")
    
  
    #armazenamento da data de hoje, no caso 08/10/2023
    
    hoje_today = dt.date.today()

    #procedimento para poder configurar o gmail com o corpo do email
    host = "smtp.gmail.com"
    port = "587"
    login = "antonyofelipe1999@gmail.com"
    senha = "pomh tydq jmbl ibvf"

    server = smtplib.SMTP(host,port)
    server.ehlo()
    server.starttls()
    server.login(login,senha)

    #não precisa do <p> e <p/> pode dar espaço vai exatamenta a mesma coisa
    corpo_email = f"""

    Prezados, ótimo trabalho a todos!   

    Seguem as cotações atualizadas das moedas de Dólar, Euro e Bitcoin, extraídas diretamente de uma API (https://economia.awesomeapi.com.br/last/USD-BRL,EUR-BRL,BTC-BRL)



    Antonio Felipe de Araujo
    (11)953473732
    """


    email_msg = MIMEMultipart()
    email_msg['From'] = login
    email_msg['To'] = login
    email_msg['Subject'] = f"Atualização das coteções USD/EUR/BIT do dia {hoje_today}"
    #plain é padrão, tem como mandar por html mas o <p> provavelmente vai contar
    email_msg.attach(MIMEText(corpo_email, 'plain'))
    
    #como anexar arquivo
    #troquei as // \\, tem que abrir o arquivo de forma binaria
    caminho_anexo = "C:/Users/Felipe/Cotacao.xlsx"
    attachment = open(caminho_anexo, "rb")
    att = MIMEBase('application', 'octet-stream')
    att.set_payload(attachment.read())
    encoders.encode_base64(att)

    att.add_header('Content-Disposition', f'attachment; filename=Cotacao.xlsx')
    attachment.close()

    email_msg.attach(att)



    #enviar o e-mail
    server.sendmail(email_msg['From'],email_msg['To'],email_msg.as_string())
    server.quit

   
    
    

    
    time.sleep(15)

 
