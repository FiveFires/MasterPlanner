from excel.datafiller import DataFiller
from excel.excel_data_manager import ExcelDataManager

if __name__ == "__main__":
    source_file_path = input("Source file path: ")
    source_sheet_name = input("Source sheet name: ")
    source_sheet_column_name_row = int(input("which row contains column names in source sheet?: ")) - 1
    destination_file_path = input("Destination file path: ")
    destination_sheet_name = input("Destination sheet name: ")
    destination_sheet_column_name_row = int(input("which row contains column names in destination sheet?: ")) - 1

    source_sheet = ExcelDataManager(source_file_path, source_sheet_name, source_sheet_column_name_row)
    destination_sheet =  ExcelDataManager(destination_file_path, destination_sheet_name, destination_sheet_column_name_row)
    
    source_sheet.read_excel()
    destination_sheet.read_excel()

    datafiller_instance = DataFiller(source_sheet, destination_sheet)

    datafiller_instance.fill_data()
