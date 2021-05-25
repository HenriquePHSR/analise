import os
import pandas as pd
import re
import openpyxl
import numpy as np
from datetime import datetime
import spreadsheet_process as sp
import unicodedata


UF = pd.read_csv('UF.csv')
ddd = pd.read_csv('ddd.csv')
UF_dict = dict(zip(list(UF['estado']), list(UF['sigla'])))
ddd_dict = dict(zip(list(ddd['ddd']), list(ddd['UF'])))

def format_string(to_format):
    return ''.join(l  for l in str(to_format) if l.isalnum() or l.isspace())

def is_email_valid(email):
    if type(email) == str:
        if re.search(r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)", email):
            return True
        else:
            return False
    else:
        return False
    
def rm_non_numeric(num):
    num = str(num)
    num = re.sub("[^0-9]",'', num)
    return num

def format_tel(tel):
    tel = str(tel)
    if len(tel) == 11 and int(tel[:2]) in ddd_dict:
        return '({}){}-{}'.format(tel[:2], tel[2:7], tel[7:])
    else:
        return tel


def rm_encode_err_terms(string):
    string = str(string)
    string = string.encode("ascii", "ignore")
    return string.decode("ascii", "ignore")

def check_UF(row, UF_dict):
    uf = str(row['Estado'])
    tels = [row['Telefone 1'], row['Telefone 2'], row['Telefone 3']]
    if uf in UF_dict.values():
        return uf
    else:
        if uf == '' or uf == None: # search in phone number
            for phone in tels:
                if phone != None and len(phone)==14 and int(phone[:2]) in ddd_dict.keys():
                    print(UF_dict[ddd_dict[int(phone[:2])]])
                    return UF_dict[ddd_dict[int(phone[:2])]]
            return None
        if len(uf)>2:
            if uf in UF_dict.keys():
                return UF_dict[uf]
            else: # search in phone number
                for phone in tels:
                    if phone != None and len(phone)==14 and int(phone[:2]) in ddd_dict.keys():
                        print(UF_dict[ddd_dict[int(phone[:2])]])
                        return UF_dict[ddd_dict[int(phone[:2])]]
                return None

def build_df(args):
    if (len(args)!=0):
        df = pd.read_excel(args[0])
    else:
        print('data path: ')
        f_name = input()
        df = pd.read_excel(f_name)


    
    df = df.drop(columns=['#','Último Acesso','Matrícula'])

    forbidden_words = ['AAA','TESTE', 'TEST', 'TRTRIXX', 'TRIXX', 'Pedro', 'HIKARO']

    # remove linhas que contenham no email ou no nome os termos acima
    df = df[~df['E-Mail'].str.upper().str.contains('|'.join(forbidden_words), na=False)]
    df = df[~df['Nome'].str.upper().str.contains('|'.join(forbidden_words), na=False)]
    df = df[~df['QLB'].str.upper().str.contains('|'.join(forbidden_words), na=False)]
    ''' AUTOCOMPLETE CPF  AND FORMAT'''
    df['CPF'] = df['CPF'].str.replace('*', '', regex=False)
    df['CPF'] = df['CPF'].str.zfill(11)

    colunas_padronizadas = ['Nome','QLB','Endereço', 'Curso', 'Complemento', 'Bairro', 'Cidade', 'Estado']

    #colunas em camelcase
    df[colunas_padronizadas] = df[colunas_padronizadas].apply(lambda x: x.str.title())
    #remov caracteres com missing type
    df[colunas_padronizadas] = df[colunas_padronizadas].apply(lambda x: x.str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8'))

    #COLUNA: NOME
    #remove linhas 
    df = df[df['Nome'].notna()]
    # df[df['Nome'].str.contains('454', na=False)]
    #remove caracteres especiais
    df['Nome'] = df['Nome'].apply(format_string)
    # COLUNA: QLB
    df['QLB'] = df['QLB'].apply(format_string)
    # DATAS
    df['Nascimento'] = df['Nascimento'].replace('0000-00-00', np.nan)
    df['Nascimento'] = pd.to_datetime(df['Nascimento'], format='%Y-%m-%d')
    df['Registro'] = pd.to_datetime(df['Registro'], format='%d/%m/%Y')
    df['Inscrição'] = pd.to_datetime(df['Inscrição'], format='%d/%m/%Y')

    df['Sexo'] = df['Sexo'].str.lower().map({'masculino': 'M', 'feminino': 'F'})

    columns = ['ID', 'Número', 'CEP', 'RG', 'Telefone 1', 'Telefone 2', 'Telefone 3']
    df[columns] = df[columns].replace(r'[^0-9]+', '',regex=True)
    df['Bairro'] = df['Bairro'].apply(lambda x: rm_encode_err_terms(x))
    for index, row in df.iterrows():
        df.loc[index, 'Estado'] = check_UF(row, UF_dict)
    
    print(df)
    print('file save: '+args[0][:-5]+'-processed.xlsx')
    df.to_excel(args[0][:-5]+'-processed.xlsx')