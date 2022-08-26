# -*- coding: utf-8 -*-
"""
Created on Mon Sep 20 16:47:46 2021

@author: jujum
"""

#colocar assinatura também
#COLOCAR TODOS OS BOLETOS NA PASTA BOLETOS A ENVIAR


#https://www.codeforests.com/2020/06/05/how-to-send-email-from-outlook/

import os
import win32com.client as win32
import pandas as pd
import numpy as np
from numpy import nan

path = "C:\\Users\\jujum\\Onedrive\\Desktop\\boletos_a_enviar"
os.chdir(path)
os.getcwd()

df = pd.read_excel('C:\\Users\\jujum\\Onedrive\\Desktop\\excel_maladireta.xlsx')
print(df)
df['email_to']

#CÓDIGO PARA LOOP FOR EM TODA A PLANLIHA EXCEL MALA DIRETA
for i in range(len(df)): 
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    email_principal = df.iloc[i,3]
    mail.To = email_principal
    mail.Subject = "Instituto da Criança | Agradecemos por exercitar a solidariedade conosco!"
    mail.GetInspector 
    nome_doador=df.iloc[i,2]
    mail.HtmlBody = """
 <html>
  <head></head>
  <body>
  	<font color="Black" size=-1 face="Arial">
        </p> %s <br>
        <br>
       Em nome do Instituto da Criança e de todas as instituições apoiadas por nós, agradecemos por você exercer a solidariedade conosco!<br>
    <br>
       Sua doação é muito importante para articularmos ações sociais.<br>
       <br>
       Anexo enviamos o boleto bancário, referente a sua contribuição.<br>
       <br>
       Agradecemos a parceria!
    </p>
    </font>
  </body>
</html>
""" % nome_doador
    arquivo = path+"\\"+df.iloc[i,0]
    mail.Attachments.Add(arquivo)
    #emai cc
    emailcc1=df.iloc[i,4]
    emailcc2=df.iloc[i,5]
    emailcc3=df.iloc[i,6]
    emailcc4=df.iloc[i,7]
    emailcc5=df.iloc[i,8]
    emailcc6=df.iloc[i,9]
    emailcc7=df.iloc[i,10]
    emailcc8=df.iloc[i,11]
    emailcc9=df.iloc[i,12]
    emailcc10=df.iloc[i,13]
    addr_cc   = [emailcc1, emailcc2,emailcc3,emailcc4,emailcc5,emailcc6,emailcc7,emailcc8,emailcc9,emailcc10]
    new_addr_cc = [item for item in addr_cc if not(pd.isnull(item)) == True]
    mail.CC="; ".join(new_addr_cc)
    #COLOCAR ADM@ SEMPRE COMO CC)
    mail.Display(False)
    mail.Send()



#######################################################################################
#outra opção
def Emailer(text, subject, recipient):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.HtmlBody = text
    mail.Display(False)
    mail.Send()
Emailer('helloooooooooo' , 'Wow it works for real' , 'jujumesquita7@yahoo.com.br')