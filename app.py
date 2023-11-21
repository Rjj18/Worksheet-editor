import pandas as pd
import openpyxl as openpyxl
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from leitura_planilha import ler_arquivo_excel




def listar_valores_coluna(dataframe, nome_coluna):
    """
    Função para listar os valores de uma coluna em um DataFrame.

    Parameters:
    - dataframe: O DataFrame do Pandas.
    - nome_coluna: O nome da coluna a ser listada.

    Returns:
    - list: Uma lista contendo os valores da coluna especificada.
    """
    try:
        # Verifica se o objeto passado é um DataFrame
        if not isinstance(dataframe, pd.DataFrame):
            raise ValueError("O objeto não é um DataFrame do Pandas.")
        
        # Verifica se a coluna existe no DataFrame
        if nome_coluna in dataframe.columns:
            # Seleciona a coluna e converte para uma lista
            valores_coluna = dataframe[nome_coluna].tolist()
            
            # Retorna a lista de valores
            return valores_coluna
        else:
            print(f"A coluna '{nome_coluna}' não foi encontrada no DataFrame.")
            return None
    
    except Exception as e:
        # Em caso de erro, imprime uma mensagem e retorna None
        print(f"Erro ao listar os valores da coluna: {e}")
        return None

def filtrar_por_valor(dataframe, nome_coluna, valor_filtro):
    """
    Função para filtrar um DataFrame com base em um valor específico em uma coluna.

    Parameters:
    - dataframe: O DataFrame do Pandas.
    - nome_coluna: O nome da coluna a ser usada para o filtro.
    - valor_filtro: O valor a ser usado como critério de filtro.

    Returns:
    - DataFrame: Um novo DataFrame contendo apenas as linhas que atendem ao critério de filtro.
    """
    try:
        # Verifica se o objeto passado é um DataFrame
        if not isinstance(dataframe, pd.DataFrame):
            raise ValueError("O objeto não é um DataFrame do Pandas.")

        # Verifica se a coluna existe no DataFrame
        if nome_coluna in dataframe.columns:
            # Filtra o DataFrame com base no valor especificado na coluna
            dataframe_filtrado = dataframe[dataframe[nome_coluna] == valor_filtro]

            # Retorna o DataFrame filtrado
            return dataframe_filtrado
        else:
            print(f"A coluna '{nome_coluna}' não foi encontrada no DataFrame.")
            return None

    except Exception as e:
        # Em caso de erro, imprime uma mensagem e retorna None
        print(f"Erro ao filtrar o DataFrame: {e}")
        return None

def remover_duplicatas(valores_coluna):
    """
    Função para remover duplicatas de uma lista de valores.

    Parameters:
    - valores_coluna: A lista de valores.

    Returns:
    - list: Uma nova lista contendo valores únicos, preservando a ordem de aparição.
    """
    try:
        # Cria um conjunto (set) para remover duplicatas e converte de volta para lista
        valores_unicos = list(set(valores_coluna))
        
        # Preserva a ordem de aparição usando sorted
        valores_unicos = sorted(valores_unicos, key=valores_coluna.index)
        
        return valores_unicos

    except Exception as e:
        # Em caso de erro, imprime uma mensagem e retorna None
        print(f"Erro ao remover duplicatas: {e}")
        return None

def gerar_dataframes_por_valor(dataframe, nome_coluna):
    """
    Função para gerar DataFrames filtrados para cada valor único em uma coluna.

    Parameters:
    - dataframe: O DataFrame do Pandas.
    - nome_coluna: O nome da coluna a ser usada para a geração dos DataFrames.

    Returns:
    - dict: Um dicionário onde as chaves são os valores únicos e os valores são DataFrames filtrados.
    """
    try:
        # Verifica se o objeto passado é um DataFrame
        if not isinstance(dataframe, pd.DataFrame):
            raise ValueError("O objeto não é um DataFrame do Pandas.")

        # Verifica se a coluna existe no DataFrame
        if nome_coluna not in dataframe.columns:
            print(f"A coluna '{nome_coluna}' não foi encontrada no DataFrame.")
            return None

        # Obtém os valores únicos da coluna
        valores_unicos = dataframe[nome_coluna].unique()

        # Dicionário para armazenar os DataFrames filtrados
        dataframes_por_valor = {}

        # Loop sobre cada valor único e gera um DataFrame filtrado
        for valor in valores_unicos:
            dataframe_filtrado = dataframe[dataframe[nome_coluna] == valor]
            dataframes_por_valor[valor] = dataframe_filtrado

        return dataframes_por_valor

    except Exception as e:
        # Em caso de erro, imprime uma mensagem e retorna None
        print(f"Erro ao gerar DataFrames por valor: {e}")
        return None

def gerar_arquivos_excel(dataframes_por_valor, sufixo='Paraisopolis'):
    """
    Função para gerar arquivos Excel para cada DataFrame em um dicionário.

    Parameters:
    - dataframes_por_valor: Um dicionário onde as chaves são os valores únicos e os valores são DataFrames.
    - sufixo: Um sufixo opcional para os nomes dos arquivos Excel.

    Returns:
    - dict: Um dicionário onde as chaves são os valores únicos e os valores são os caminhos dos arquivos Excel gerados.
    """
    try:
        # Dicionário para armazenar os caminhos dos arquivos Excel gerados
        caminhos_arquivos = {}

        # Loop sobre cada valor único e gera um arquivo Excel
        for valor, df in dataframes_por_valor.items():
            # Cria um novo livro do Excel
            workbook = Workbook()

            # Adiciona o DataFrame ao livro do Excel
            sheet = workbook.active
            for row in dataframe_to_rows(df, index=False, header=True):
                sheet.append(row)

            # Define o caminho do arquivo
            caminho_arquivo = f"/workspace/gerenciadorPlanilhas/planilhas-filtradas/{valor} {sufixo}.xlsx"

            # Salva o arquivo Excel
            workbook.save(caminho_arquivo)

            # Adiciona o caminho do arquivo ao dicionário
            caminhos_arquivos[valor] = caminho_arquivo

        return caminhos_arquivos

    except Exception as e:
        # Em caso de erro, imprime uma mensagem e retorna None
        print(f"Erro ao gerar arquivos Excel: {e}")
        return None


caminho_arquivo = "/workspace/gerenciadorPlanilhas/lista-nominal_PARAISOPOLIS.xlsx"
dados_excel = ler_arquivo_excel(caminho_arquivo)
coluna_selecionada = 'TP_UNIDADE'  # Substitua pelo nome da coluna desejada
valores_da_coluna = listar_valores_coluna(dados_excel, coluna_selecionada)
valores_unicos = remover_duplicatas(valores_da_coluna)
dataframes_por_valor = gerar_dataframes_por_valor(dados_excel, coluna_selecionada)
arquivos_gerados = gerar_arquivos_excel(dataframes_por_valor)




# Verifica se a operação foi bem-sucedida
if arquivos_gerados is not None:
    # Exibe os caminhos dos arquivos Excel gerados
    for valor, caminho_arquivo in arquivos_gerados.items():
        print(f"Arquivo gerado para o valor '{valor}': {caminho_arquivo}")
else:
    print("Falha ao gerar arquivos Excel.")