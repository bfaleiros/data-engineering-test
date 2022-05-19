from openpyxl import load_workbook
import pandas as pd
from datetime import datetime


def capturabasepivot(nmpivot):

    # Criação da coleção que armazenará a base de dados final reestruturada
    dct_final = {
        "year_month": [],
        'uf': [],
        'product': [],
        'unit': [],
        'volume': [],
        'created_at': []
    }

    df_final = pd.DataFrame(dct_final) # Dataframe Pandas para armazenamento da coleção de dados da tabela fato final e reestruturada


    # Verifica se a Pivot Table selecionada está entre as solicitadas no teste
    if nmpivot != 'Tabela dinâmica1' and nmpivot != 'Tabela dinâmica3':
        print('Favor selecionar entre as tabelas dinâmicas 1 (Tabela dinâmica1) ou 3 (Tabela dinâmica3)')
        return df_final


    # Carregamento inicial da planilha
    myarq = load_workbook('vendas-combustiveis-m3.xlsx').active


    # Obtém os índices das Pivot Tables existentes na planilha
    lst_pivottables = []
    for nmpvt in myarq._pivots:
        lst_pivottables.append(nmpvt.name)


    # Captura o cache da Pivot Table selecionada
    mypivot = myarq._pivots[lst_pivottables.index(nmpivot)]


    # Criação das listas utilizadas no processo
    lst_produto = []  # Lista contendo a relação de produtos da base de dados
    lst_uf = []  # Lista contendo a relação de estados da base de dados
    lst_ano = []  # Lista contendo a relação de anos da base de dados
    lst_transacional = []  # Lista que armazenará 1 registro completo da base de dados


    # Criação da coleção que armazenará a base de dados original (cache) completa
    dct_fato = {
        "PRODUTO": [],
        'ANO': [],
        'REGIAO': [],
        'ESTADO': [],
        'UNIDADE': [],
        'Jan': [],
        'Fev': [],
        'Mar': [],
        'Abr': [],
        'Mai': [],
        'Jun': [],
        'Jul': [],
        'Ago': [],
        'Set': [],
        'Out': [],
        'Nov': [],
        'Dez': [],
        'TOTAL': []
    }

    df_fato = pd.DataFrame(dct_fato) # Dataframe Pandas para armazenamento da coleção de dados da tabela fato (base original em cache)


    ###
    # Rotina de iteração com a base cache da Pivot Table selecionada para obtenção
    # das listas indexadas das dimensões (Produto, Ano e Estado)
    ###
    for myfields in mypivot.cache.cacheFields:
        for myfieldsvalue in myfields.sharedItems._fields:
            if myfields.name == 'COMBUSTÍVEL':
                try:  # Tratamento de falha para valores não encontrados
                    lst_produto.append(myfieldsvalue.v)
                except AttributeError:
                    lst_produto.append(None)
            elif myfields.name == 'ANO':
                try:  # Tratamento de falha para valores não encontrados
                    lst_ano.append(int(myfieldsvalue.v))
                except AttributeError:
                    lst_ano.append(None)
            elif myfields.name == 'ESTADO':
                try:  # Tratamento de falha para valores não encontrados
                    lst_uf.append(myfieldsvalue.v)
                except AttributeError:
                    lst_uf.append(None)


    ###
    # Rotina de iteração com a base cache transacional para captura
    # dos registros da base de dados da Pivot Table selecionada
    ###
    for mr in mypivot.cache.records.r:
        lst_transacional = []
        for mf in mr._fields:
            try:  # Tratamento de falha para valores não encontrados
                lst_transacional.append(mf.v)
            except AttributeError:
                lst_transacional.append(0)
        df_fato.loc[len(df_fato)] = lst_transacional


    # Lista indexada de meses para utilização posterior no processo de normalização da base final
    lst_meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']


    # Rotina de iteração com a coleção de dados obtida no cache
    # criada para normalizar e rotular os campos de dimensões

    lst_final = []

    for index, row in df_fato.iterrows():
        for i in range(len(lst_meses)):
            lst_final.append(datetime(lst_ano[row['ANO']], i+1, 1))
            lst_final.append(lst_uf[row['ESTADO']])
            lst_final.append(lst_produto[row['PRODUTO']])
            lst_final.append(row['UNIDADE'])
            lst_final.append(row[f'{lst_meses[i]}'])
            lst_final.append(datetime.now())
            df_final.loc[len(df_final)] = lst_final
            lst_final = []

    del myarq  # Limpa cache de memória obtido no processo

    return df_final

# Chamada e captura da base de dados da Pivot Table selecionada
# Selecionar entre as Pivot Tables Tabela dinâmica1 e Tabela dinâmica3

my_PivotDB = capturabasepivot('Tabela dinâmica1')

if not my_PivotDB.empty:
    print(my_PivotDB[['year_month', 'uf', 'product', 'unit', 'volume', 'created_at']])
    print(my_PivotDB.info())


