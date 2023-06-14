"""
Module: ExcelDataManager
Description: This module provides a class for reading, writing, 
                and appending data to an Excel file using pandas.

Author: Adam Ondryas
Email: adam.ondryas@gmail.com

This software is distributed under the GPL v3.0 license.
"""

# external module imports
import pandas as pd

class ExcelDataManager():
    def __init__(self, file_path, sheet_name=0, column_name_row=0):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.column_name_row = column_name_row
        self.df = None

    def read_excel(self):
        try:
            return (pd.read_excel(self.file_path, self.sheet_name, skiprows=self.column_name_row))
        
        except FileNotFoundError:
            print(f"File '{self.file_path}' not found.")
        
        except Exception as e:
            print(f"An Error occured while reading the Excel file: {str(e)}")

    def write_excel(self, index=False):
        try:
            writer = pd.ExcelWriter(self.file_path, sheet_name=self.sheet_name, engine='openpyxl')
            self.df.to_excel(writer, sheet_name=self.sheet_name, index=index)
            writer.save()
            print(f"Data successfully written to '{self.file_path}'.")

        except Exception as e:
            print(f"An error occurred while writing the Excel file: {str(e)}")

    def append_to_excel(self, data, index=False, startcol=0):
        try:
            with pd.ExcelWriter(self.file_path, mode='a', engine='openpyxl', if_sheet_exists="overlay") as writer:
                data.to_excel(writer, sheet_name=self.sheet_name, index=index,
                              startrow=self.column_name_row, startcol=startcol)
                print(f"Data successfully appended to '{self.file_path}'.")
                
        except Exception as e:
            print(f"An error occured while appending to the Excel file: {str(e)}")

