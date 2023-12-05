import pandas as pd
import openpyxl as openpyxl

def read_excel_file(caminho_arquivo):

//Parameters: - excel file path.
    try:
        dataframe = pd.read_excel(caminho_arquivo)
        return dataframe
    
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return None
