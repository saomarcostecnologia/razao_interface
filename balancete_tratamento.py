import time
import pandas as pd
import re
from datetime import date
from datetime import datetime
from openpyxl import Workbook
import numpy as np
import datetime

def tratar_balancete(file_path):
    #Leitura de dataframe e manipulação de planilha

    balancete = pd.read_excel(file_path, header=None)

    balancetemoeda = balancete

    #subir linha 12 para cabeçalho

    # Armazenando a primeira linha como cabeçalho
    novo_cabecalho = balancete.iloc[12]

    #subir o cabeçalho para coluna.
    balancete.columns = novo_cabecalho

    # Resetando os índices
    balancete = balancete.reset_index(drop=True)
    #Remoção das primeiras linhas.

    balancete = balancete.drop(balancete.index[:13])

    balancete = balancete.dropna(subset=["CONTA"])

    #numero de conta de niveis

    #balancete['concat_2'] = balancete['CONTA'].str.split(('-')).str[0] 
    balancete['Nivel'] = balancete['CONTA'].str.split(('                      ')).str[0].str.split().str[0]
    balancete['id'] = balancete['CONTA'].str.split(('-')).str[0].str.split().str[0]

    #Tratamento de DataFrame para conecatenar linhas vazias.

    tratamento01 = pd.DataFrame({'id': balancete['id'],'valor': balancete['Nivel']})

    tratamento01 = tratamento01[tratamento01['valor'].isna()]

    tratamento01['Setimo_nivel'] = balancete['id']

    #Mesclagem da base balancete e tratamento01

    balancete_tratado = pd.merge(balancete, tratamento01, left_on= 'id', right_on= 'id', how='left')

    #Preenchimento de colunas vazias de Nivel

    balancete_tratado['Nivel'].fillna(method='ffill', inplace=True)

    #Concatenando colunas.

    balancete_tratado['Nivel_Correto'] = balancete_tratado['Nivel'].astype(str) + balancete_tratado['Setimo_nivel'].astype(str)

    #removendo duplicidade em coluna conta.

    balancete_tratado = balancete_tratado.drop_duplicates(subset='Nivel_Correto', keep='last')

    balancete_tratado['Nivel_Correto'] = balancete_tratado['Nivel_Correto'].str.split(('nan')).str[0]

    #Organizando Balancete

    #removendo colunas.

    balancete_tratado = balancete_tratado.drop(['Setimo_nivel','valor','id','Nivel'], axis=1)

    #Mudando nome de colunas

    balancete_tratado.rename(columns={'Nivel_Correto': 'Nivel'}, inplace=True)

    balancete_tratado.head(10)

    nome_coluna = ['CONTA', 'SALDO ANTERIOR','C/D1','DEBITO','CREDITO','SALDO ATUAL','C/D2','Nivel']

    balancete_tratado.columns = nome_coluna

    # organizando data frame

    balancete_tratado =  balancete_tratado[['Nivel','CONTA', 'SALDO ANTERIOR','C/D1','DEBITO','CREDITO','SALDO ATUAL','C/D2']]

    # Resetar o índice

    balancete_tratado = balancete_tratado.reset_index(drop=True)

    balancete_etl2 = balancete_tratado.tail(5)

    # Adicione uma nova coluna 'Nova_Coluna' com valores nulos (None)
    balancete_etl2.loc[:, 'Nova_Coluna'] = None

    #Organizando o dataframe

    balancete_etl2 =  balancete_etl2[['Nivel','CONTA','SALDO ANTERIOR','Nova_Coluna','C/D1','DEBITO','CREDITO','SALDO ATUAL','C/D2',]]

    # Realiza as renomeações de colunas
    balancete_etl2.rename(columns={'C/D1': 'DEBITO', 'DEBITO': 'CREDITO', 'CREDITO': 'SALDO ATUAL', 'SALDO ATUAL': 'C/D1'}, inplace=True)

    nova_ordem = balancete_etl2[['Nivel', 'CONTA', 'SALDO ANTERIOR', 'C/D1', 'DEBITO', 'CREDITO', 'SALDO ATUAL', 'C/D2']]
    balancete_etl2 = nova_ordem


    #remoção das ultimas linhas do balancete

    balancete_tratado = balancete_tratado.drop(balancete_tratado.tail(5).index)

    # Assuming you have a DataFrame called balancete_tratado already defined
    dados = {
        'Nivel': [None],
        'CONTA': [None],
        'SALDO ANTERIOR': [None],
        'C/D1': [None],
        'DEBITO': [None],
        'CREDITO': [None],
        'SALDO ATUAL': [None],
        'C/D2': [None],
    }

    nova_linha = pd.DataFrame(dados)

    # Concatenating the new row to the original DataFrame
    balancete_tratado = pd.concat([balancete_tratado, nova_linha], ignore_index=True)

    # Print the DataFrame
    balancete_tratado.tail(5)

    # Concatena o dataframe 'balancete_etl_01' no final do dataframe 'balancete_tratado'
    balancete_tratado = pd.concat([balancete_tratado, balancete_etl2], ignore_index=True)

    #balancete_tratado.to_excel('b.xlsx', index=False)

    # Criar um nome dinâmico para o arquivo tratado
    # Remover os últimos 5 caracteres da variável file_path (assumindo que sejam ".xlsx")
    file_path_sem_extensao = file_path[:-5]
    hora_atual = datetime.datetime.now().strftime("%H%M%S")  # Obtém a hora atual no formato HHMMSS
    nome_arquivo_tratado = f"{file_path_sem_extensao}.{hora_atual}.xlsx"
    # Salvar o arquivo tratado com o nome dinâmico
    balancete_tratado.to_excel(nome_arquivo_tratado, sheet_name='Tratado', index=False)

    return nome_arquivo_tratado
    #razao.to_excel('teste10.xlsx',sheet_name='Tratado',index=False)