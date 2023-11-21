import pandas as pd
import openpyxl as openpyxl

def ler_arquivo_excel(caminho_arquivo):
    """
    Função para ler um arquivo Excel usando o Pandas.

    Parameters:
    - caminho_arquivo (str): O caminho do arquivo Excel.

    Returns:
    """
    try:
        # Utiliza a função read_excel do Pandas para ler o arquivo Excel
        dataframe = pd.read_excel(caminho_arquivo)
        
        # Retorna o DataFrame resultante
        return dataframe
    
    except Exception as e:
        # Em caso de erro, imprime uma mensagem e retorna None
        print(f"Erro ao ler o arquivo Excel: {e}")
        return None