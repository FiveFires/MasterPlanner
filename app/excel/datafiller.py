"""
Module: DataFiller
Description: This module provides a class for filling data in a destination Excel sheet
             based on common lookup values between a source Excel sheet and the destination sheet.
             
Author: Adam Ondryas
Email: adam.ondryas@gmail.com
"""

# local module imports
from .excel_data_manager import ExcelDataManager

# other module imports
import pandas as pd
import numpy as np

class DataFiller():
    def __init__(self, source: ExcelDataManager, destination: ExcelDataManager):
        self.source = source
        self.destination = destination
        print("SOURCE SHEET'S -->")
        self.source.lookup_range = input("Lookup column name or lookup value: ")
        self.source.column_name = input("Name of the column that you want to copy from: ")

        print("DESTINATION SHEET'S -->")
        self.destination.column_name = input("Name of the column that you want to fill: ")

    def fill_data(self):
        # Getting unique lookup values from source and destination dataframes
        source_unique_lookup_values = self.__get_unique_lookup_values(self.source.df)
        destination_unique_lookup_values = self.__get_unique_lookup_values(self.destination.df)

        # Finding common values between source and destination lookup values
        unique_lookup_values = self.__find_common_values(destination_unique_lookup_values, 
                                                        source_unique_lookup_values)
        
        # Iterating over common lookup values
        for current_lookup_value in unique_lookup_values:
            # Creating subsets of dataframes based on current lookup value
            current_source_subset = self.source.df[
                (self.source.df[self.source.lookup_range] == current_lookup_value)]
            
            current_destination_subset = self.destination.df[
                (self.destination.df[self.source.lookup_range] == current_lookup_value)]
        
            # Aligning indices of destination subset with source subset
            self.__align_indeces(current_destination_subset, current_source_subset)
            
            # Updating values in destination column with values from source column
            self.__update_values(self.destination.df[self.destination.column_name], current_source_subset[self.source.column_name])

        # Appending updated dataframe to the destination Excel file
        self.destination.append_to_excel(self.destination.df)

        
    def __get_unique_lookup_values(self, df):
        return df[self.source.lookup_range].unique()
    
    def __find_common_values(self, array1, array2):
        return np.intersect1d(array1, array2)
        
    def __update_values(self, df_to_be_updated, df_with_new_values):
        df_to_be_updated.update(df_with_new_values)

    def __align_indeces(self, base_df, to_update_df):
        to_update_df.index = base_df.index[0:(to_update_df.shape[0])]
