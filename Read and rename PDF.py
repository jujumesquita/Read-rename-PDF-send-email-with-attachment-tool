# -*- coding: utf-8 -*-
"""
Created on Sun Sep 19 14:50:50 2021

@author: jujum
"""

#FERRAMENTA PARA LER O PDF E CRIAR UM DATAFRAME COM AS INFORMAÇÕES SOBRE OS BOLETOS

#INÍCIO DO CÓDIGO


import os

path = "C:\\Users\\jujum\\Onedrive\\Desktop\\boletos_extrair_info"
ref_boleto="05.2022"


os.chdir(path)
os.getcwd()

#Listar arquivos de uma pasta
from os import listdir
from os.path import isfile, join
onlyfiles = [f for f in listdir(path) if isfile(join(path, f))]
onlyfiles

#Converter lista "onlyfiles" para dataframe
import pandas as pd 
df = pd.DataFrame(onlyfiles) 
df.columns = ['nome_arquivos']
print(df) 

#Adicionar colunas ao df
df.insert(1, "valor_dt_vcto", "") 
df.insert(2, "dt_vcto", "") 
df.insert(3, "nome_doador_end", "") 
df.insert(4, "nome_doador_final", "") 
df.insert(5, "valor_int", "") 
df.insert(6, "valor_cent", "") 
df.insert(7, "valor_final", "") 
df.insert(8, "mes_ano_ref", "") 
df.insert(9, "nome_doador1", "") 
df.insert(10, "nome_doador2", "") 

print(df) 

#Print texto de todos os arquivos pdf na pasta e extração das infos
import PyPDF2
for i in range(len(df)):    
    arquivo = df.iloc[i,0]
    with open(arquivo, mode='rb') as f:
        reader = PyPDF2.PdfFileReader(f)
        page = reader.getPage(0)
        #print(page.extractText())
        texto=page.extractText()
        split1 = (texto.split('237-2'))
        parte2=split1[2]
        parte1 = split1[0]
        print(parte1)  #ver como partir o valor da data de vencimento
        df.iat[i,1]=parte1
        split2 = (parte2.split('DMR$'))
        parte3=split2[1]
        split3=(parte3.split(','))
        parte4=split3[0]
        df.iat[i,3]=parte4
        parte5=(parte4.split(" "))
        nome_doadorfinal=parte5[0]
        df.iat[i,4]=nome_doadorfinal
        split5=(parte1.split(','))
        df.iat[i,5]=split5[0]
        df.iat[i,9]=parte5[0]
        df.iat[i,10]=parte5[1]
        df.iat[i,6]=split5[1]
 
#Extraindo data de vencimento        
df['dt_vcto']=df['valor_dt_vcto'].str[-10:]

#Extraindo centavos
df['valor_cent'] = df['valor_cent'].str[:2]


#Formando valor final     
df['valor_final']=df.valor_int+","+df.valor_cent
df.to_excel("output.xlsx")

#Colocando nome 1a maiuscula e outras minusculas
df['nome_doador_final']=df['nome_doador_final'].apply(lambda x: x.lower())
df['nome_doador_final']=df['nome_doador_final'].apply(lambda x: x.title())

print(df)


#Concatenar nome+referência em nova coluna
df.mes_ano_ref=ref_boleto
df['nome_arquivo_novo']= df.nome_doador1+" "+df.nome_doador2+" "+"- REF"+" "+df.mes_ano_ref+".pdf"

df.to_excel("output.xlsx")


#CHECAR NOMES DOS ARQUIVOS NOVOS NO ARQUIVO OUTPUT PARA VER SE ESTÃO OK E CONSERTAR QUALQUER ERRO ERRO

#LER ARQUIVO OUTPUT ATUALIZADO PARA PEGAR NOMES CORRETOS DOS ARQUIVOS
os.getcwd()
df = pd.read_excel('output.xlsx')
print(df)

#Mudar o nome do arquivo com o nome do novo arquivo
os.getcwd()
for i in range(len(df)):    
    arquivo = df.iloc[i,1]
    old_file = arquivo
    new_file = df.iloc[i,12]
    os.rename(old_file, new_file)


