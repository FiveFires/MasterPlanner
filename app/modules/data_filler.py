"""
Module: DataFiller
Description: This module provides a class for filling data in a destination Excel sheet
             based on common lookup values between a source Excel sheet and the destination sheet.
             
Author: Adam Ondryas
Email: adam.ondryas@gmail.com

This software is distributed under the GPL v3.0 license.
"""

# local module imports
from .excel_data_manager import ExcelDataManager

# external module imports
import pandas as pd
import numpy as np

class DataFiller():
    def __init__(self, source: ExcelDataManager, destination: ExcelDataManager,
                 src_lookup_column, src_copy_column, 
                 dst_lookup_column, dst_fill_column):
        
        self.source = source
        self.source.lookup_column = src_lookup_column
        self.source.copy_column = src_copy_column

        self.destination = destination
        self.destination.lookup_column = dst_lookup_column
        self.destination.fill_column = dst_fill_column

    def fill_data(self, progress_callback=None):
        # Getting unique lookup values from source and destination dataframes
        source_unique_lookup_values = self._get_unique_lookup_values(self.source.df, self.source.lookup_column)
        destination_unique_lookup_values = self._get_unique_lookup_values(self.destination.df, self.destination.lookup_column)

        # Finding common values between source and destination lookup values
        unique_lookup_values = self._find_common_values(destination_unique_lookup_values, 
                                                        source_unique_lookup_values)
        
        num_of_unique_lookup_values = unique_lookup_values.size

        # Iterating over common lookup values
        for index, current_lookup_value in enumerate(unique_lookup_values):
            # Creating subsets of dataframes based on current lookup value
            current_source_subset = self.source.df[
                (self.source.df[self.source.lookup_column] == current_lookup_value)]
            
            current_destination_subset = self.destination.df[
                (self.destination.df[self.destination.lookup_column] == current_lookup_value)]
        
            # Aligning indices of destination subset with source subset
            self._align_indeces(current_destination_subset, current_source_subset)
            
            # Updating values in destination column with values from source column
            self._update_values(self.destination.df[self.destination.fill_column], current_source_subset[self.source.copy_column])
            
            progress_callback.emit(int( (index / (num_of_unique_lookup_values-1)) * 100 ))

        # Get the index of column to be appended
        fill_column_index = self.destination.df.columns.get_loc(self.destination.fill_column)

        # Appending updated dataframe to the destination Excel file
        self.destination.append_to_excel(self.destination.df[self.destination.fill_column], 
                                         startcol=fill_column_index)
        
    def _get_unique_lookup_values(self, df, lookup_column):
        return df[lookup_column].unique()
    
    def _find_common_values(self, array1, array2):
        return np.intersect1d(array1, array2)
        
    def _update_values(self, df_to_be_updated, df_with_new_values):
        df_to_be_updated.update(df_with_new_values)

    def _align_indeces(self, base_df, to_update_df):
        row_diff = base_df.shape[0] - to_update_df.shape[0]

        if(row_diff > 0):
            to_update_df.index = base_df.index[0:(to_update_df.shape[0])]

        elif(row_diff < 0):
            to_update_df.drop(to_update_df.tail(abs(row_diff)).index, inplace=True)
            to_update_df.index = base_df.index

        elif(base_df.shape[0] == to_update_df.shape[0]):
            to_update_df.index = base_df.index