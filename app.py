import pandas as pd
import openpyxl as openpyxl
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from excel_reader import read_excel_file

def extract_column_values(dataframe, column_name):
    try:
        if not isinstance(dataframe, pd.DataFrame):
            raise ValueError("Invalid input. Expected a Pandas DataFrame.")

        if column_name not in dataframe.columns:
            print(f"Column '{column_name}' not found in the DataFrame.")
            return None

        return dataframe[column_name].tolist()

    except Exception as e:
        print(f"Error extracting column values: {e}")
        return None

def filter_dataframe(dataframe, column_name, filter_value):
    try:
        if not isinstance(dataframe, pd.DataFrame):
            raise ValueError("Invalid input. Expected a Pandas DataFrame.")

        if column_name not in dataframe.columns:
            print(f"Column '{column_name}' not found in the DataFrame.")
            return None

        return dataframe[dataframe[column_name] == filter_value]

    except Exception as e:
        print(f"Error filtering the DataFrame: {e}")
        return None

def remove_duplicates(values):
    try:
        unique_values = list(set(values))
        return sorted(unique_values, key=values.index)

    except Exception as e:
        print(f"Error removing duplicates: {e}")
        return None

def generate_dataframes_by_value(dataframe, column_name):
    try:
        if not isinstance(dataframe, pd.DataFrame):
            raise ValueError("Invalid input. Expected a Pandas DataFrame.")

        if column_name not in dataframe.columns:
            print(f"Column '{column_name}' not found in the DataFrame.")
            return None

        values = dataframe[column_name].unique()
        return {value: dataframe[dataframe[column_name] == value] for value in values}

    except Exception as e:
        print(f"Error generating DataFrames by value: {e}")
        return None

def generate_excel_files(dataframes_by_value, suffix='Paraisopolis'):
    try:
        file_paths = {}

        for value, df in dataframes_by_value.items():
            workbook = Workbook()
            sheet = workbook.active

            for row in dataframe_to_rows(df, index=False, header=True):
                sheet.append(row)

            file_path = f"/workspace/spreadsheetManager/filtered-spreadsheets/{value} {suffix}.xlsx"
            workbook.save(file_path)
            file_paths[value] = file_path

        return file_paths

    except Exception as e:
        print(f"Error generating Excel files: {e}")
        return None

# Specify the path to the Excel file
file_path = "/workspace/spreadsheetManager/nominal-list_PARAISOPOLIS.xlsx"

# Read the Excel file into a Pandas DataFrame
excel_data = read_excel_file(file_path)

# Specify the column to work with
selected_column = 'UNIT_TYPE'  # Replace with the desired column name

# List values in the selected column
column_values = extract_column_values(excel_data, selected_column)

# Remove duplicates from the list of column values
unique_values = remove_duplicates(column_values)

# Generate DataFrames for each unique value in the selected column
dataframes_by_value = generate_dataframes_by_value(excel_data, selected_column)

# Generate Excel files for each DataFrame
generated_files = generate_excel_files(dataframes_by_value)

# Check if the operation was successful
if generated_files:
    for value, file_path in generated_files.items():
        print(f"File generated for '{value}': {file_path}")
else:
    print("Failed to generate Excel files.")
