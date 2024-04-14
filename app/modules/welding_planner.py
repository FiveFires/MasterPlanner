"""
Module: WeldingPlanner
Description: This module contains the welding planner tool. TBD
             
Author: Adam Ondryas
Email: adam.ondryas@gmail.com

This software is distributed under the GPL v3.0 license.
"""

# local module imports
from .excel_data_manager import ExcelDataManager

# external module imports
from decimal import ROUND_UP
import pandas as pd
import numpy as np
from datetime import date
from isoweek import Week
import os

class PlannerMX():
    def __init__(self, mx):
        self.mx = mx
        self.name = None
        self.inventory = 0
        self.df = pd.DataFrame(columns=["reservation", "project", "deadline"])
        self.temp_pieces_in_batch = 0
        self.batch_size = 0
        self.cooperation_time = 0

class WeldingPlanner():
    def __init__(self, welding_planner_excel=None):
        self.welding_planner_excel = welding_planner_excel
        self.planner_mx = None
        self.production_batches = []
        self.batch_database_missing_parts = []

    def plan_welding(self, bi_reservations_excel, manufacturing_plan_excel, 
                     batch_database_excel, progress_callback=None):

        # Read and load data from the input Excel files
        bi_reservations_excel.df = bi_reservations_excel.read_excel()
        self._update_progress_bar(progress_callback, 5)
        
        manufacturing_plan_excel.df = manufacturing_plan_excel.read_excel()
        self._update_progress_bar(progress_callback, 10)

        batch_database_excel.df = batch_database_excel.read_excel()
        self._update_progress_bar(progress_callback, 15)

        if self.welding_planner_excel != None:
            self.welding_planner_excel.df = self.welding_planner_excel.read_excel()

        # Fill all the NaNs to zero in STAV_MAT column
        self._fill_empty_cells(bi_reservations_excel.df["STAV_MAT"], 0)

        # Get a list of unique material numbers
        unique_MXs = self._get_unique_values_in_column(bi_reservations_excel.df, "CISLO_MAT")

        # Iterate through all unique material numbers
        for index, current_mx in enumerate(unique_MXs):
            self._update_progress_bar(progress_callback, int( (((index+1) /  unique_MXs.size) * 80) + 20) )

            # Create a PlannerMX object for the current material number
            self.planner_mx = PlannerMX(current_mx)

            # Retrieve batch size and manufacturing time for the material number
            self._get_batch_and_manufacturing_time(batch_database_excel.df)

            if(self.planner_mx.batch_size == 0):
                # Skip if the batch size is zero (missing parts in the batch database)
                self.batch_database_missing_parts.append(self.planner_mx.mx)
                continue

            # Filter and fill the MX planner with reservations data
            self._filter_fill_mx_planner(bi_reservations_excel.df)
            
            # Skip if the dataframe is empty or if the inventory is sufficient
            if( (self.planner_mx.df.shape[0] == 0) or (self._is_inventory_sufficient()) ):
                continue

            # Retrieve project deadlines from the manufacturing plan dataframe
            self._get_project_deadlines(manufacturing_plan_excel.df)
            
            if(self.planner_mx.df.shape[0] == 0):
                continue

            # Drop projects covered by the inventory from the MX planner dataframe
            self._drop_projects_covered_by_inventory()

            # Generate production batches based on the MX planner dataframe
            self._generate_production_batches()

        # Generate the output Excel file
        self._generate_output_excel()

    def _fill_empty_cells(self, df, fill_value):
        df.fillna(fill_value, inplace=True)
        
    def _get_unique_values_in_column(self, df, column_title):
        return (df[column_title].unique())
    
    def _is_inventory_sufficient(self):
        return (self.planner_mx.inventory >= int(self.planner_mx.df["reservation"].sum()) )
    
    def _filter_fill_mx_planner(self, reservations_df):
        # Filter rows based on specific conditions from the reservations dataframe
        filtered_rows = reservations_df[(reservations_df['CISLO_MAT'] == self.planner_mx.mx) &
                                        (reservations_df['_IB_KOKS'] != "S2024") & 
                                        (reservations_df['CIS_OBJ'].str.get(1) != "K") &
                                        (reservations_df['CIS_OBJ'] != 0)]

        # Further filter rows to remove duplicates and specific values from "_IB_KOKS" column
        filtered_rows = filtered_rows[((~filtered_rows.duplicated("_IB_KOKS")) | 
                                       (filtered_rows["_IB_KOKS"] == "M2023")  | 
                                       (filtered_rows["_IB_KOKS"] == "M2024") )]
        
        if(filtered_rows.shape[0] >= 1):
            filtered_rows.reset_index(drop=True)

            # Fill PlannerMX object with the filtered data
            self.planner_mx.name = filtered_rows["NAZEV_MAT"].values[0]
            self.planner_mx.inventory = int(filtered_rows["STAV_MAT"].values[0]) + self._get_count_in_manufacturing()

            self.planner_mx.df["reservation"] = filtered_rows["MNOZSTVI"]
            self.planner_mx.df["project"] = filtered_rows["_IB_KOKS"]
            self.planner_mx.df["deadline"] = filtered_rows["DODATUMU"].dt.isocalendar().week

    def _get_count_in_manufacturing(self):
        retval = 0

        if(self.welding_planner_excel is not None):
            # Filter rows based on material number and batches in production
            in_manufacturing = self.welding_planner_excel.df[( (self.welding_planner_excel.df["MATERIAL NUMBER"] == self.planner_mx.mx) &
                                                               (self.welding_planner_excel.df["BATCH IN PRODUCTION"].values > 0) )]

            if(in_manufacturing.shape[0] >= 1):
                    # Calculate the sum of items in production batches
                    retval = int(in_manufacturing["BATCH IN PRODUCTION"].sum())

        return retval
    
    def _get_project_deadlines(self, manufacturing_plan_df):
        # Filter out M2022 and M2023 projects from the PlannerMX dataframe
        filtered_planner_mx_df = self.planner_mx.df[(self.planner_mx.df["project"] != "M2023") & (self.planner_mx.df["project"] != "M2024")]

        # Get unique project numbers from the filtered PlannerMX dataframe
        unique_project_numbers = self._get_unique_values_in_column(filtered_planner_mx_df, "project")

        # Filter the manufacturing plan dataframe based on unique project numbers, "Unnamed: 9" column is the project numbers column
        filtered_manufacturing_plan_df = manufacturing_plan_df.loc[manufacturing_plan_df["Unnamed: 9"].isin(unique_project_numbers)]

        if(filtered_planner_mx_df.shape[0] != filtered_manufacturing_plan_df.shape[0]):
            # Drop duplicates in the manufacturing plan dataframe based on project numbers
            filtered_manufacturing_plan_df = filtered_manufacturing_plan_df.drop_duplicates(subset="Unnamed: 9")

            if(filtered_planner_mx_df.shape[0] != filtered_manufacturing_plan_df.shape[0]):
                # Find projects that were found in the manufacturing plan
                projects_found = filtered_manufacturing_plan_df["Unnamed: 9"]

                # Filter the PlannerMX dataframe based on found projects
                filtered_planner_mx_df = filtered_planner_mx_df[filtered_planner_mx_df["project"].isin(projects_found)]

                self.planner_mx.df = self.planner_mx.df[self.planner_mx.df["project"].isin(projects_found) |
                                                       (self.planner_mx.df["project"] == "M2023") |
                                                       (self.planner_mx.df["project"] == "M2024")]

        # Sort the dataframes by project number and index them to match
        filtered_planner_mx_df = filtered_planner_mx_df.sort_values("project", ascending=True)
        filtered_manufacturing_plan_df = filtered_manufacturing_plan_df.sort_values("Unnamed: 9", ascending=True)
        filtered_manufacturing_plan_df.index = filtered_planner_mx_df.index

        # Update the PlannerMX dataframe's "deadline" column with the manufacturing plan's delivery week
        self.planner_mx.df.loc[(self.planner_mx.df["project"] != "M2023") & 
                               (self.planner_mx.df["project"] != "M2024"), "deadline"] = filtered_manufacturing_plan_df["CURRENT DELIVERY WEEK "]
        
    def _drop_projects_covered_by_inventory(self):
        # Sort the merged subsets by delivery week and reset row indeces
        self.planner_mx.df = self.planner_mx.df.sort_values('deadline', ascending=True)
        self.planner_mx.df = self.planner_mx.df.reset_index(drop=True)

        # Iterate over the count of reserved pieces and continually subtract them from the inventory count
        # remove a row each time until inventory count gets to zero -->
        # --> this selects only the projects that do not have enough materials for them
        for reserved_pieces_count in self.planner_mx.df['reservation'].values:
            self.planner_mx.inventory -= reserved_pieces_count
            
            # Project is covered by inventory, can drop the row
            if(self.planner_mx.inventory >= 0):
                self.planner_mx.df = self.planner_mx.df.drop(self.planner_mx.df.index[0])

            # Project is not covered by inventory, break the cycle
            # and add the leftover inventory count to the first batch count
            elif(self.planner_mx.inventory < 0):
                self.temp_pieces_in_batch = (self.planner_mx.inventory + reserved_pieces_count)
                break

    def _get_batch_and_manufacturing_time(self, batch_databse_df):
            filtered_batch_database_df = batch_databse_df[batch_databse_df["Číslo"] == self.planner_mx.mx]

            # Check if not empty
            if(filtered_batch_database_df.shape[0] != 0):
                manufacturing_cooperation_time = filtered_batch_database_df["Norma Kooperace"].values[0]

                # Check if the material cooperation time exists, if == "X", it does not exist yet, skip this CurrentMaterialNumber
                if(manufacturing_cooperation_time != "X"):
                    self.planner_mx.cooperation_time = int(np.ceil(manufacturing_cooperation_time/7))
                    self.planner_mx.batch_size = int(filtered_batch_database_df["Dávka"].values[0])

    def _generate_production_batches(self):
        TodaysDate = date.today()

        # CONFIG VARIABLES --> CAN BE CHANGED TO BE MORE OR LESS CONSERVATIVE
        #MaterialDeliveryTimeInWeeks = 4
        MaterialPickingTimeInWeeks = 2
        AssemblyTimeInWeeks = 1

        production_batch = [0] * 6
        
        while(self.planner_mx.df.shape[0] > 0):
            # Set the values for this production batch
            production_batch[0] = self.planner_mx.mx
            production_batch[1] = self.planner_mx.name
            production_batch[2] = self.planner_mx.batch_size
            try:                        
                production_batch[3] = Week(TodaysDate.year, int(self.planner_mx.df['deadline'].values[0]
                                                    - MaterialPickingTimeInWeeks 
                                                    - AssemblyTimeInWeeks)).monday()
            except:
                print(f"failed at {self.planner_mx.mx}")

            production_batch[4] = Week(TodaysDate.year, int(self.planner_mx.df['deadline'].values[0]
                                                    - MaterialPickingTimeInWeeks
                                                    - AssemblyTimeInWeeks
                                                    - self.planner_mx.cooperation_time)).monday()
            production_batch[5] = 0

            # Add the production batch to the list of batches
            self.production_batches.append(production_batch.copy())

            # Initialize temporary variables
            self.planner_mx.temp_pieces_in_batch += self.planner_mx.batch_size
            num_of_rows_to_be_deleted = 0

            # Figure out how many projects (rows) are covered by each production batch
            for reserved_pieces_count in self.planner_mx.df['reservation'].values:                      
                self.planner_mx.temp_pieces_in_batch -= reserved_pieces_count
                
                # If the batch covers more than one reservation, continue subtracting
                if(self.planner_mx.temp_pieces_in_batch > 0):
                    num_of_rows_to_be_deleted += 1
                
                # If the reservations deplete the batch, drop the covered rows and break to create a new batch
                elif(self.planner_mx.temp_pieces_in_batch == 0):
                    num_of_rows_to_be_deleted += 1
                    self.planner_mx.df.drop(self.planner_mx.df.index[0:num_of_rows_to_be_deleted], inplace=True)
                    break

                # If the batch has excess pieces after covering reservations, adjust and drop the covered rows
                elif(self.planner_mx.temp_pieces_in_batch < 0):
                    self.planner_mx.temp_pieces_in_batch += reserved_pieces_count
                    self.planner_mx.df.drop(self.planner_mx.df.index[0:num_of_rows_to_be_deleted], inplace=True)
                    break

                # If the number of rows to be deleted exceeds the available rows, drop all the rows
                if(num_of_rows_to_be_deleted >= self.planner_mx.df.shape[0]):
                    self.planner_mx.df.drop(self.planner_mx.df.index[0:num_of_rows_to_be_deleted], inplace=True)
                    break

    def _generate_output_excel(self):
        # Create a DataFrame from the production batches:
        welding_plan_df = pd.DataFrame(self.production_batches, 
                                       columns=["MATERIAL NUMBER", 
                                                "NAME", 
                                                "PIECES IN BATCH", 
                                                "READY FOR PICKING", 
                                                "WELDING COMPLETED", 
                                                "BATCH IN PRODUCTION"])
        
        # Sort the DataFrame by "READY FOR PICKING" column in ascending order
        welding_plan_df.sort_values("READY FOR PICKING", ascending=True, inplace=True)

        # Reset the index of the DataFrame
        welding_plan_df.reset_index(drop=True, inplace=True)

        if(self.welding_planner_excel != None):
            # Filter rows from the welding planner Excel DataFrame where "BATCH IN PRODUCTION" is greater than 0
            in_manufacturing_rows = self.welding_planner_excel.df[self.welding_planner_excel.df["BATCH IN PRODUCTION"].values > 0]
            in_manufacturing_rows.reset_index(drop=True, inplace=True)

            # Concatenate the filtered rows with the welding plan DataFrame
            welding_plan_df = pd.concat([in_manufacturing_rows, welding_plan_df], ignore_index=True)

        # Specify the output path for the Excel file
        path = "./output"

        # Check whether the specified path exists or not
        isExist = os.path.exists(path)

        if not isExist:
            # Create a new directory if it does not exist
            os.makedirs(path)

        # Create an Excel file and write the DataFrames to different sheets
        with pd.ExcelWriter("output/WeldingPlan.xlsx") as writer:
            # Write the welding plan DataFrame to the "WeldingPlan" sheet
            welding_plan_df.to_excel(writer, sheet_name="Welding Plan", index=False)

            if(len(self.batch_database_missing_parts) > 0):
                # Create a DataFrame from the batch database missing parts list
                x_database_missing_df = pd.DataFrame(self.batch_database_missing_parts, columns=["MATERIAL NUMBER"])

                # Write the batch database missing parts DataFrame to the "X_database missing" sheet
                x_database_missing_df.to_excel(writer, sheet_name="X_database missing", index=False)

    def _update_progress_bar(self, progress_callback, percentage):
        try:
            progress_callback.emit(percentage)
        except Exception as e:
            print("An error occured while updating the progress bar")

