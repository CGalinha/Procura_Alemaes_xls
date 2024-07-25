"""
nova versao do arquivo Procura_xls.py
data : 25/06/2024


A funcionar 25/06/2024

"""


import os
import pandas as pd
import datetime


class Bcolors:
    OK = "\033[92m"  # GREEN
    WARNING = "\033[93m"  # YELLOW
    FAIL = "\033[91m"  # RED
    RESET = "\033[0m"  # RESET COLOR

def fich_xls():
    tmp = []
    df = pd.read_excel("busca.xlsx")
    #print(df)
    busca = df['Pesquisa'].tolist()
    for a in busca:
        # print(type(a))
        if type(a) is int:
            b = str(a)
            tmp.append(b)
        else:
            tmp.append(a)
    # print(tmp)
    return tmp

def exp1(palavra_procurada):
    print(f'Lista da procura -> {palavra_procurada}')
    # Defina a palavra que você está procurando

    # Resolvido com um ficheiro em ecxel
    # palavra_procurada = ['T470s', 'T460', 'T480s', 'T490s',
    # 'T580', 'Yoga', 'Carbon', 'Inspiron 13-5378', 'M920qs', '7370','7380','(AIO)','7460', ' G8', 'G9','iPad','Apple']

    # Pasta onde os arquivos do Excel estão localizados
    pasta_excel = "down/"

    # Loop através dos arquivos na pasta

    for arquivo in os.listdir(pasta_excel):
        if arquivo.endswith(".xlsx") or arquivo.endswith(".xls"):

            # Caminho completo para o arquivo
            caminho_arquivo = os.path.join(pasta_excel, arquivo)
            print(f'o Caminho do ficheiro a trabalhar -> {caminho_arquivo} ')

            # Leitura do arquivo Excel
            df = pd.read_excel(caminho_arquivo)

            # print(f'{Bcolors.OK} O xls a ser trabalhado {Bcolors.RESET}')

            nome = f'escolha_{arquivo}'
            # print(f'{Bcolors.FAIL}nome do ficheiro a ser gravado ->{Bcolors.RESET} - {nome}')

            # Vai a procura
            procura(df, palavra_procurada, nome)
    ################################################

def procura(df, procura, nome):
    cont = 0
    # print(query)
    resultados = [] # Criar lista vazia para armazenar as linhas correspondentes
    # Criar um DataFrame vazio para armazenar as linhas correspondentes

    for busca in procura:
        cont = 0
        print(f'Esta a procura de ->{busca}')

    # Iterar sobre as linhas do DataFrame original
        for indice, linha in df.iterrows():
            # Verificar se a palavra está na linha
            if busca in str(linha.values):
                # print(f'Encontrou a palavra {busca}')
                # Copiar a linha para o novo DataFrame
                cont = cont + 1
                # print(linha)
                print(f'      ---->  encontrou Linha -> {busca}')
                resultados.append(linha)

    if not resultados:
        print('   ------ Sem resultados ---------')
    else:
        # cria uma  Variavel datatime com a data e hora
        
        data_hora_atual = datetime.datetime.now()
        
        
        df_resultados = pd.DataFrame(resultados)
        # print('   ------ os resultados foram ---------')
        nome = f'final/final_{data_hora_atual}.xlsx'
        # Salva os resultados em um novo arquivo Excel
        # print(nome)
        df_resultados.to_excel(nome, index=False)
        print(f'Para a chave {busca} encontrou.se {cont} de linhas ')
        print(f'{Bcolors.FAIL}nome do ficheiro a ser gravado ->{Bcolors.RESET} - {nome}')
    print('##########################   FIM   #####################################')


if __name__ == "__main__":
    busca = fich_xls()
    exp1(busca)