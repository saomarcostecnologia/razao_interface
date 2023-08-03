import time
import pandas as pd
import re
from datetime import date
from datetime import datetime
import datetime


#Leitura de dataframe e manipulação de planilha
def tratar_razao(file_path):
    
    # Leitura de dataframe e manipulação de planilha
    razao = pd.read_excel(file_path, header=None)
    
    diretorio = razao
    razao_cabecalho = diretorio
    razao = diretorio
    razao = razao.drop(razao.index[:11])

    #subir linha 15 para cabeçalho

    # Armazenando a primeira linha como cabeçalho
    novo_cabecalho = razao_cabecalho.iloc[15]

    # Definindo a primeira linha como cabeçalho
    razao = razao[1:]
    razao.columns = novo_cabecalho

    # Resetando os índices
    razao = razao.reset_index(drop=True)

    #df para a segunda parte do tratamento
    razao_pt2 = razao

    #Coluna para tratamento

    etl_contas = razao[['DIA','UA']]

    etl_contas = pd.DataFrame(etl_contas, columns=['DIA','UA'])

    etl_contas_2 = etl_contas

    etl_contas_2 = etl_contas_2.drop(etl_contas_2[(etl_contas_2 == 0).any(axis=1) |
                                                etl_contas_2.isnull().any(axis=1) |
                                                etl_contas_2['UA'].str.startswith('UA') |
                                                etl_contas_2['DIA'].str.startswith('CONTAS ANTERI') |
                                                etl_contas_2['UA'].str.startswith('PROP') 
                                                ]
                                                .index)
    etl_contas_2 = etl_contas_2.reset_index(drop=True)

    # Criar uma lista com a sequência "banco" e "estrutura"
    sequencia = ["1","2","3"]
    # Atribuir a sequência à coluna "ColunaNova" do DataFrame
    etl_contas_2["ColunaNova"] = sequencia * (len(etl_contas_2) // len(sequencia)) + sequencia[:len(etl_contas_2) % len(sequencia)]

    # Realizar o pivô
    etl_contas_3 = etl_contas_2.pivot(columns='ColunaNova', values='UA')

    #Preenchimento de colunas vazias.

    etl_contas_3['1'].fillna(method='ffill', inplace=True)
    etl_contas_3['2'].fillna(method='ffill', inplace=True)

    #Exclusão de vazios na coluna 3
    etl_contas_3 = etl_contas_3.dropna(subset=["3"])

    #Reset index
    etl_contas_3 = etl_contas_3.reset_index(drop=True)


    #Seleção de colunas
    etl_contas_3 = etl_contas_3[['2','3']]

    etl_contas_3 = pd.DataFrame(etl_contas_3, columns=['2','3'])

    #Mudando nome de Coluna

    etl_contas_3.rename(columns={'2': 'Estrutura de Contas'}, inplace=True)
    etl_contas_3.rename(columns={'3': 'Conta'}, inplace=True)

    estrutura = etl_contas_3

    # Segunda parte do Tratamento

    #Base INICIAL COM CABEÇALHO TRATADO
    razao = razao_pt2

    #Mesclagem de Estrutura de conta + razão

    razao = pd.merge(razao, estrutura, left_on= 'UA', right_on= 'Estrutura de Contas', how='left')

    #Preenchimento de colunas vazias de Estrtura de Contas e Contas

    razao['Estrutura de Contas'].fillna(method='ffill', inplace=True)
    razao['Conta'].fillna(method='ffill', inplace=True)

    #removendo o periodo da coluna 0
    razao = razao.drop(razao.index[:6])

    # Resetando os índices
    razao = razao.reset_index(drop=True)

    razao['Nova Coluna'] = razao['DIA']
    razao['Nova Coluna'] = razao['DIA'].str[:5]

    ### Nome da coluna que contém as informações
    nome_coluna = 'Nova Coluna'

    # Nome específico para identificar o início do range
    nome_inicio_range = 'MOVTO'

    # Nome específico para identificar o fim do range (que contém "PERÍODO")
    nome_fim_range = 'PERÍO'

    # Identificando o índice da linha de início do range
    indice_inicio_range = razao[razao[nome_coluna] == nome_inicio_range].index[0]
    indice_fim_range = razao[razao[nome_coluna] == nome_fim_range].index[0]

    x = len(razao['DIA'])

    # Nome específico para identificar o início do range
    nome_inicio_range = 'MOVTO'

    # Nome específico para identificar o fim do range (que contém "PERÍODO")
    nome_fim_range = 'PERÍO'

    # Inicializando os índices de início e fim
    indice_inicio_range = 0
    indice_fim_range = 0

    # Loop para remover todos os ranges de linhas
    while True:
        # Verificando se há um novo range a ser removido
        if nome_inicio_range in razao[nome_coluna].values and nome_fim_range in razao[nome_coluna].values:
            # Identificando o índice da linha de início do range
            indice_inicio_range = razao[razao[nome_coluna] == nome_inicio_range].index[0]
            # Identificando o índice da linha de fim do range
            indice_fim_range = razao[razao[nome_coluna] == nome_fim_range].index[0]
            
            # Removendo o range de linhas
            razao = razao.drop(range(indice_inicio_range, indice_fim_range + 1))
            # Resetando os índices
            razao = razao.reset_index(drop=True)
        else:
            # Se não há mais ranges a serem removidos, sair do loop
            break

    #TERCEIRA PARTE DO TRATAMENTO

    # Definir o tamanho do conjunto de linhas
    tamanho_conjunto_linhas = 4

    # Inicializar o índice de início do conjunto de linhas
    indice_inicio_conjunto = 0

    #Colunas que serao armazenadas
    coluna1 = []
    coluna2 = []
    coluna3 = []
    coluna4 = []


    while indice_inicio_conjunto < len(razao):
        # Selecionar o conjunto de linhas atual
        linhas_selecionadas = razao['LOTE'].iloc[indice_inicio_conjunto:indice_inicio_conjunto+tamanho_conjunto_linhas]
        
        # Verificar se há linhas suficientes para formar o conjunto completo
        if len(linhas_selecionadas) == tamanho_conjunto_linhas:
            # Extrair as informações das linhas selecionadas para novas colunas
            df_extracted = linhas_selecionadas.str.extract(r'(\w+.*)')
            
            # Transpor o DataFrame resultante
            df_transposed = df_extracted.transpose()
            
            # Resetar o índice
            df_transposed = df_transposed.reset_index(drop=True)
            
            # Adicionar as informações nas listas correspondentes
            coluna1.extend(df_transposed.iloc[:, 0])
            coluna2.extend(df_transposed.iloc[:, 1])
            coluna3.extend(df_transposed.iloc[:, 2])
            coluna4.extend(df_transposed.iloc[:, 3])
    
        # Atualizar o índice de início do próximo conjunto de linhas
        indice_inicio_conjunto += tamanho_conjunto_linhas

    # Remover linhas vazias da coluna "dia" da planilha razao
    razao = razao.dropna(subset=["DIA"])

    # Resetar o índice
    razao = razao.reset_index(drop=True)

    # Criar um novo DataFrame com as colunas extraídas
    df_extraido = pd.DataFrame({
        'Coluna1': coluna1,
        'Coluna2': coluna2,
        'Coluna3': coluna3,
        'Coluna4': coluna4
    })

    # Concatenar o DataFrame extraído com a planilha original
    razao_concatenada = pd.concat([razao, df_extraido], axis=1)

    # Excluir as 3 últimas linhas da planilha razao
    razao_concatenada = razao_concatenada.iloc[:-3]

    # Resetar o índice
    razao_concatenada = razao_concatenada.reset_index(drop=True)

    razao = razao_concatenada

    razao.rename(columns={'Coluna1': 'Origem'}, inplace=True)
    razao.rename(columns={'Coluna2': 'Categoria'}, inplace=True)
    razao.rename(columns={'Coluna3': 'Lote'}, inplace=True)
    razao.rename(columns={'Coluna4': 'Moeda'}, inplace=True)
    razao = razao.drop(['Nova Coluna','LOTE'], axis=1)
    razao = razao.dropna(axis=1, how='all')

    #Quarta parte do Tratamento

    #Mudando tipo do dado.
    razao['CREDITO'] = razao['CREDITO'].astype(float)
    razao['DEBITO'] = razao['DEBITO'].astype(float)

    #substituindo valores null para 0

    razao['CREDITO']  = razao['CREDITO'] .fillna(0)
    razao['DEBITO']  = razao['DEBITO'] .fillna(0)

    # CONCATENANDO VALORES.

    razao['valores'] = - razao['CREDITO']   + razao['DEBITO']  

    #Removendo colunas

    razao = razao.drop(['CREDITO','DEBITO'], axis=1)

    #Quinta parte do tratamento

    #Remoção dos nomes antes do " : "
    razao['Origem'] = razao['Origem'].str.replace('.*:', '', regex=True)
    razao['Categoria'] = razao['Categoria'].str.replace('.*:', '', regex=True)
    razao['Moeda'] = razao['Moeda'].str.replace('.*:', '', regex=True)


    # Remove tudo antes do caractere ':' na coluna 'Origem'
    razao['Origem'] = razao['Origem'].str.replace('.*:', '', regex=True)

    # Remove tudo antes do caractere ':' na coluna 'Categoria'
    razao['Categoria'] = razao['Categoria'].str.replace('.*:', '', regex=True)

    # Remove tudo antes do caractere ':' na coluna 'Moeda'
    razao['Moeda'] = razao['Moeda'].str.replace('.*:', '', regex=True)

    # Substitui o caractere '.' por '/' na coluna 'DIA'
    razao['DIA'] = razao['DIA'].str.replace('.', '/')

    # Obtém a data de hoje
    data_hoje = date.today()

    # Obtém o ano da data de hoje
    ano = data_hoje.year

    # Converte o ano para string
    ano = str(ano)

    # Concatena o dia do DataFrame com o ano para formar 'dia/mes/ano'
    razao['dia_mes_ano']  = razao['DIA'] + '/' + ano

    # Remove espaços em branco da coluna 'dia_mes_ano'
    razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace(' ', '')

    # Substitui ex. 'JUN' por ex.'06' na coluna 'dia_mes_ano'
    razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('JAN', '01')
    razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('FEV', '02')
    razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('MAR', '03')
    razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('ABR', '04')
    razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('MAI', '05')
    razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('JUN', '06')
    razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('JUL', '07')
    razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('AGO', '08')
    razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('SET', '09')
    razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('OUT', '10')
    razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('NOV', '11')
    razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('DEZ', '12')

    # Converte a coluna 'dia_mes_ano' para o tipo de dado datetime usando o formato "%d/%m/%Y"
    razao['Data'] = pd.to_datetime(razao['dia_mes_ano'], format="%d/%m/%Y")

    # Remove as colunas 'dia_mes_ano' e 'DIA' do DataFrame
    razao = razao.drop(['dia_mes_ano','DIA'], axis=1)

    # Obtém a coluna 'Data' tratada
    data_tratada = razao['Data']

    # Converte a coluna 'Data' para o tipo de dado datetime e extrai apenas a data, sem a informação do horário
    data_tratada = pd.to_datetime(data_tratada, format="%d/%m/%Y").dt.date

    # Substitui a coluna 'Data' pelo resultado do tratamento
    razao['Data'] = data_tratada

    # organizando data frame

    nova_ordem = ['Estrutura de Contas','Conta','Data','HISTORICO','Lote','valores','UA','C/C','CONTRAP','Origem','Categoria','Moeda']

    razao = razao[nova_ordem]

    # Criar um nome dinâmico para o arquivo tratado
    file_path_sem_extensao = file_path[:-5]
    hora_atual = datetime.datetime.now().strftime("%H%M%S")  # Obtém a hora atual no formato HHMMSS
    nome_arquivo_tratado = f"{file_path_sem_extensao}.{hora_atual}.xlsx" 
    # Salvar o arquivo tratado com o nome dinâmico
    razao.to_excel(nome_arquivo_tratado, sheet_name='Tratado', index=False)

    return nome_arquivo_tratado
    #razao.to_excel('teste10.xlsx',sheet_name='Tratado',index=False)
