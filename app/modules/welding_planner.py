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

class WeldingPlanner():
    def __init__(self, bi_reservations_excel: ExcelDataManager,
                 manufacturing_plan_excel: ExcelDataManager,
                 batch_database_excel: ExcelDataManager,
                 welding_planner_excel=None):
        
        self.bi_reservations_excel = bi_reservations_excel
        self.manufacturing_plan_excel = manufacturing_plan_excel
        self.batch_database_excel = batch_database_excel
        self.welding_planner_excel = welding_planner_excel
        
    def plan_welding(self, progress_callback=None):
        Sheet_ListOfReservations = self.bi_reservations_excel.df
        Sheet_ManufacturingPlan = self.manufacturing_plan_excel.df
        Sheet_ManufacturingTime = self.batch_database_excel.df

        if self.welding_planner_excel != None:
            Sheet_WeldingPlan8000 = self.welding_planner_excel.df
            ResetPlan = False
        else:
            ResetPlan = True

        TodaysDate = date.today()
        ManufacturingTimeNotDefined_Count = 0
        OutputList = [0] * 6
        FinalOutputList = []
        XDatabaseMissingPartsList = []

        # CONFIG VARIABLES --> CAN BE CHANGED TO BE MORE OR LESS CONSERVATIVE
        #MaterialDeliveryTimeInWeeks = 4
        MaterialPickingTimeInWeeks = 2
        AssemblyTimeInWeeks = 1

        # Fill all the NaNs to zero in STAV_MAT column
        Sheet_ListOfReservations.fillna(0, inplace=True)
        # Get a list of Unique material numbers
        UniqueMaterialNumbers = Sheet_ListOfReservations['CISLO_MAT'].unique()

        UniqueMaterialNumberCount = UniqueMaterialNumbers.size

        # Iterate through all Unique material numbers
        for index, CurrentMaterialNumber in enumerate(UniqueMaterialNumbers):
            try:
                progress_callback.emit(int((( (index+1)/UniqueMaterialNumberCount) * 80) +20))
            except:
                pass
            # Get a subset table with all the current material numbers without the rows that have project number set as S2023
            CurrentMaterialNumber_Entries = Sheet_ListOfReservations[(Sheet_ListOfReservations['CISLO_MAT'] == CurrentMaterialNumber) & 
                                                                    (Sheet_ListOfReservations['_IB_KOKS'] != "S2023") & 
                                                                    (Sheet_ListOfReservations['_IB_KOKS'] != "M2023") & 
                                                                    (Sheet_ListOfReservations['CIS_OBJ'].str.get(1) != "K") &
                                                                    (Sheet_ListOfReservations['CIS_OBJ'] != 0)]
            
            # Check if CurrentMaterialNumber_Entries is not empty
            if(CurrentMaterialNumber_Entries.shape[0] >= 1):
                ### WARNING: MIGHT CAUSE ISSUES LATER FOR PROPER PLANNING, FIND BETTER SOLUTION ###
                #Drop rows with duplicit Project Numbers that were found in the Reservation List before going forward
                CurrentMaterialNumber_Entries = CurrentMaterialNumber_Entries.drop_duplicates(subset='_IB_KOKS')

                # Get a Reservations count of current material numbers
                CurrentMaterialNumber_ReservationsCount = CurrentMaterialNumber_Entries['MNOZSTVI'].sum()

                # Get an inventory count of current material numbers
                CurrentMaterialNumber_InventoryCount = int(CurrentMaterialNumber_Entries.loc[CurrentMaterialNumber_Entries.index[0],'STAV_MAT'])

                if(ResetPlan == False):
                    CurrentMaterialNumber_AlreadyInManufacturing_Entries = Sheet_WeldingPlan8000[(Sheet_WeldingPlan8000["MATERIAL NUMBER"] == CurrentMaterialNumber) &
                                                                                        (Sheet_WeldingPlan8000["BATCH IN PRODUCTION"].values > 0)]
                    
                    if(CurrentMaterialNumber_AlreadyInManufacturing_Entries.shape[0] >= 1):
                        CurrentMaterialNumber_AlreadyInManufacturing_Count = CurrentMaterialNumber_AlreadyInManufacturing_Entries["BATCH IN PRODUCTION"].sum()
                        CurrentMaterialNumber_InventoryCount += CurrentMaterialNumber_AlreadyInManufacturing_Count

                # Check if there is enough material
                if (CurrentMaterialNumber_ReservationsCount > CurrentMaterialNumber_InventoryCount):
                    # If there is not enough material, get all the current material number's unique project numbers
                    CurrentMaterialNumber_UniqueProjectNumbers = CurrentMaterialNumber_Entries['_IB_KOKS'].unique()

                    # Get the rows with all the current Unique project numbers for the current material number from the manufacturing plan sheet
                    CurrentUniqueProjectNumbers_Entries = Sheet_ManufacturingPlan.loc[Sheet_ManufacturingPlan['Unnamed: 9'].isin(CurrentMaterialNumber_UniqueProjectNumbers)]

                    # Check if some Project Numbers that are found in Manufacturing Plan are aligned with Reservation Project Numbers
                    if(CurrentUniqueProjectNumbers_Entries.shape[0] != CurrentMaterialNumber_Entries.shape[0]):

                        # Delete rows with duplicit Project Numbers from Manufacturing Plan
                        CurrentUniqueProjectNumbers_Entries = CurrentUniqueProjectNumbers_Entries.drop_duplicates(subset='Unnamed: 9')

                        # Check if Project Numbers from Manufacturing Plan are now aligned with Reservation Project Numbers
                        if(CurrentUniqueProjectNumbers_Entries.shape[0] != CurrentMaterialNumber_Entries.shape[0]):

                            # Get a list of Project Numbers that was found in the Manufacturing Plan
                            ManufacturingPlan_UniqueProjectNumbers = CurrentUniqueProjectNumbers_Entries['Unnamed: 9']

                            # Update the Current Material Number entries subset with only valid project numbers that were actually found in the Manufacturing Plan
                            CurrentMaterialNumber_Entries = CurrentMaterialNumber_Entries[CurrentMaterialNumber_Entries['_IB_KOKS'].isin(ManufacturingPlan_UniqueProjectNumbers)]

                    # PREPARE BOTH SUBSETS FOR MERGE
                    # Sort the rows in the Manufacturing Plan subset by Project Number aka unnamed 9
                    CurrentUniqueProjectNumbers_Entries = CurrentUniqueProjectNumbers_Entries.sort_values('Unnamed: 9', ascending=True)
                    # Sort the rows in the ListOfReservations subset by Project Number aka _IB_KOKS
                    CurrentMaterialNumber_Entries = CurrentMaterialNumber_Entries.sort_values('_IB_KOKS', ascending=True)

                    # Align indexes to merge correctly
                    CurrentUniqueProjectNumbers_Entries.index = CurrentMaterialNumber_Entries.index

                    # Merge the Project numbers and Delivery weeks into the Current Material Number subset sheet
                    CurrentMaterialNumber_Entries = CurrentMaterialNumber_Entries.join(CurrentUniqueProjectNumbers_Entries['CURRENT DELIVERY WEEK '])
                    CurrentMaterialNumber_Entries = CurrentMaterialNumber_Entries.join(CurrentUniqueProjectNumbers_Entries['Unnamed: 9'])

                    # Sort the merged subsets by delivery week and reset row indeces
                    CurrentMaterialNumber_Entries = CurrentMaterialNumber_Entries.sort_values('CURRENT DELIVERY WEEK ', ascending=True)
                    CurrentMaterialNumber_Entries = CurrentMaterialNumber_Entries.reset_index(drop=True)

                    # Reset the variable to not carryover batch counts from previous Material Numbers
                    Temp_PiecesInBatch = 0

                    # Iterate over the count of reserved pieces and continually subtract them from the inventory count
                    # remove a row each time until inventory count gets to zero -->
                    # --> this selects only the projects that do not have enough materials for them
                    for ReservedPiecesCount in CurrentMaterialNumber_Entries['MNOZSTVI'].values:
                        CurrentMaterialNumber_InventoryCount -= ReservedPiecesCount
                        
                        # Project is covered by inventory, can drop the row
                        if(CurrentMaterialNumber_InventoryCount >= 0):
                            CurrentMaterialNumber_Entries = CurrentMaterialNumber_Entries.drop(CurrentMaterialNumber_Entries.index[0])

                        # Project is not covered by inventory, break the cycle
                        # and add the leftover inventory count to the first batch count
                        elif(CurrentMaterialNumber_InventoryCount < 0):
                            Temp_PiecesInBatch = (CurrentMaterialNumber_InventoryCount + ReservedPiecesCount)
                            break

                    CurrentMaterialNumber_ManufacturingTimeEntries = Sheet_ManufacturingTime[Sheet_ManufacturingTime['Číslo'] == CurrentMaterialNumber]
                    if(CurrentMaterialNumber_ManufacturingTimeEntries.shape[0] == 0):
                        ManufacturingTimeNotDefined_Count += 1
                        XDatabaseMissingPartsList.append(CurrentMaterialNumber)
                    else:
                        # Get the material cooperation time cell value #
                        CurrentMaterialNumber_ManufacturingCooperationTime = CurrentMaterialNumber_ManufacturingTimeEntries['Norma Kooperace'].values[0]
                        
                        # Check if the material cooperation time exists, if == "X", it does not exist yet, skip this CurrentMaterialNumber
                        if(CurrentMaterialNumber_ManufacturingCooperationTime == "X"):
                            ManufacturingTimeNotDefined_Count += 1
                            XDatabaseMissingPartsList.append(CurrentMaterialNumber)
                        else:
                            CurrentMaterialNumber_ManufacturingCooperationTime = int(np.ceil(CurrentMaterialNumber_ManufacturingCooperationTime/7))
                            CurrentMaterialNumber_PiecesInBatch = int(CurrentMaterialNumber_ManufacturingTimeEntries['Dávka'].values[0])

                            while(CurrentMaterialNumber_Entries.shape[0] > 0):
                                # Fill the row (list)
                                OutputList[0] = CurrentMaterialNumber
                                OutputList[1] = CurrentMaterialNumber_Entries['NAZEV_MAT'].values[0]
                                OutputList[2] = CurrentMaterialNumber_PiecesInBatch                        
                                OutputList[3] = Week(TodaysDate.year, int(CurrentMaterialNumber_Entries['CURRENT DELIVERY WEEK '].values[0] - 
                                                    MaterialPickingTimeInWeeks - 
                                                    AssemblyTimeInWeeks)).monday()
                                
                                OutputList[4] = Week(TodaysDate.year, int(CurrentMaterialNumber_Entries['CURRENT DELIVERY WEEK '].values[0] - 
                                                                        MaterialPickingTimeInWeeks - 
                                                                        AssemblyTimeInWeeks - 
                                                                        CurrentMaterialNumber_ManufacturingCooperationTime)).monday()
                                OutputList[5] = 0
                                # Add the row to the list of lists
                                FinalOutputList.append(OutputList.copy())

                                # Setup temp vars
                                Temp_PiecesInBatch += CurrentMaterialNumber_PiecesInBatch
                                NumberOfRowsToBeDeleted = 0

                                # Figure out how many projects(rows) are covered by each production batch
                                for ReservedPiecesCount in CurrentMaterialNumber_Entries['MNOZSTVI'].values:                      
                                    Temp_PiecesInBatch -= ReservedPiecesCount
                                    
                                    # Batch covers more than one Reservation, continue subtracting
                                    if(Temp_PiecesInBatch > 0):
                                        NumberOfRowsToBeDeleted += 1
                                    
                                    # Reservations depleted the batch, drop the covered rows and break to make a new batch
                                    elif(Temp_PiecesInBatch == 0):
                                        NumberOfRowsToBeDeleted += 1
                                        CurrentMaterialNumber_Entries.drop(CurrentMaterialNumber_Entries.index[0:NumberOfRowsToBeDeleted], inplace=True)
                                        break

                                    elif(Temp_PiecesInBatch < 0):
                                        Temp_PiecesInBatch += ReservedPiecesCount
                                        CurrentMaterialNumber_Entries.drop(CurrentMaterialNumber_Entries.index[0:NumberOfRowsToBeDeleted], inplace=True)
                                        break

                                    if(NumberOfRowsToBeDeleted >= CurrentMaterialNumber_Entries.shape[0]):
                                        CurrentMaterialNumber_Entries.drop(CurrentMaterialNumber_Entries.index[0:NumberOfRowsToBeDeleted], inplace=True)
                                        break

        WeldingPlan = pd.DataFrame(FinalOutputList, columns=["MATERIAL NUMBER", "NAME", "PIECES IN BATCH", "READY FOR PICKING [CW]", "WELDING COMPLETED [CW]", "BATCH IN PRODUCTION"])
        WeldingPlan.sort_values("READY FOR PICKING [CW]", ascending=True, inplace=True)
        WeldingPlan.reset_index(drop=True, inplace=True)

        if(ResetPlan == False):
            WeldingPlan_InManufacturingRows = Sheet_WeldingPlan8000[Sheet_WeldingPlan8000["BATCH IN PRODUCTION"].values > 0]
            WeldingPlan_InManufacturingRows.reset_index(drop=True, inplace=True)
            WeldingPlan = pd.concat([WeldingPlan_InManufacturingRows, WeldingPlan], ignore_index=True)


        path = "./output"
        # Check whether the specified path exists or not
        isExist = os.path.exists(path)
        if not isExist:
            # Create a new directory because it does not exist
            os.makedirs(path)

        with pd.ExcelWriter("output/WeldingPlan.xlsx") as writer:
            # use to_excel function and specify the sheet_name and index
            # to store the dataframe in specified sheet
            WeldingPlan.to_excel(writer, sheet_name="WeldingPlan", index=False)
            if(ManufacturingTimeNotDefined_Count > 0):
                XDatabaseMissingParts = pd.DataFrame(XDatabaseMissingPartsList, columns=["MATERIAL NUMBER"])
                XDatabaseMissingParts.to_excel(writer, sheet_name="X_database missing", index=False)
