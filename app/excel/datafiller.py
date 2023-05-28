
#This module works similar to the VLOOKUP fnc in excel
#It must match the source sheet 

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
        self.source.lookup_range = input("Lookup column name or lookup value:")
        self.source.column_name = input("Name of the column that you want to copy from:")

        print("DESTINATION SHEET'S -->")
        self.destination.column_name = input("Name of the column that you want to fill:")

    def fill_data(self):
        source_unique_lookup_values = self.__get_unique_lookup_values(self.source.df)
        destination_unique_lookup_values = self.__get_unique_lookup_values(self.destination.df)

        unique_lookup_values = self.__find_common_values(destination_unique_lookup_values, 
                                                        source_unique_lookup_values)
        
        for current_lookup_value in unique_lookup_values:
            current_source_subset = self.source.df[
                (self.source.df[self.source.lookup_range] == current_lookup_value)]
            
            current_destination_subset = self.destination.df[
                (self.destination.df[self.source.lookup_range] == current_lookup_value)]
        
            self.__update_subset(current_destination_subset, current_source_subset)

            self.destination.append_to_excel(current_destination_subset)

        

    def __get_unique_lookup_values(self, df):
        return df[self.source.lookup_range].unique()
    
    def __find_common_values(self, array1, array2):
        return np.intersect1d(array1, array2)
        
    
    def __update_subset(self, df_to_be_updated, df_with_new_values):
        df_to_be_updated.update(df_with_new_values)
