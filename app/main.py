from excel.datafiller import DataFiller
from excel.excel_data_manager import ExcelDataManager

if __name__ == "__main__":
    source_sheet = ExcelDataManager("input data\cpr_database.xlsx")
    destination_sheet =  ExcelDataManager("input data\svarovaci_plan.xlsx", sheet_name="Svarovaci Plan")
    
    source_sheet.read_excel()
    destination_sheet.read_excel()

    datafiller_instance = DataFiller(source_sheet, destination_sheet)

    datafiller_instance.fill_data()
