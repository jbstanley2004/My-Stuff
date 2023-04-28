import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook

def read_excel_file(file_path, sheet_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {str(e)}")
        return None

def perform_excel_operation(file_path, sheet_name, operation, *args, **kwargs):
    try:
        if operation == "read":
            return read_excel_file(file_path, sheet_name)
        elif operation == "write":
            return write_to_excel_file(file_path, sheet_name, *args, **kwargs)
        else:
            print(f"Unsupported Excel operation: {operation}")
            return None
    except Exception as e:
        print(f"An error occurred while performing Excel operation: {str(e)}")
        return None

def write_to_excel_file(file_path, sheet_name, df):
    try:
        book = openpyxl.Workbook()
        writer = pd.ExcelWriter(file_path, engine='openpyxl')
        writer.book = book
        df.to_excel(writer, sheet_name=sheet_name, index=False)

        writer.save()
        writer.close()

        return True
    except Exception as e:
        print(f"An error occurred while writing to the Excel file: {str(e)}")
        return False
