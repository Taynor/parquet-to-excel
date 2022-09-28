from openpyxl import Workbook
import pandas as pd

class ParquetWorksheet:
    #fields of the class
    def __init__(self, parquet_file, excel_file, worksheets, workbook, parquet_data, parquet_filter, parquet_column, default_sheet_name):
        
        #these fields are available for the user to manipulate
        self.parquet_file = parquet_file
        self.excel_file = excel_file
        self.parquet_filter = parquet_filter
        self.default_sheet_name = default_sheet_name
        
        #these fields are hidden and are for PRIVATE use only
        self.worksheets = worksheets = []
        self.workbook = workbook
        self.parquet_data = parquet_data
        self.parquet_column = parquet_column
    
    #property get and set for parquet file value
    @property
    def parquet_file(self):
        return self.__parquet_file
    @parquet_file.setter
    def parquet_file(self, value):
        self.__parquet_file = value
    
    #property get and set for excel file value
    @property
    def excel_file(self):
        return self.__excel_file
    @excel_file.setter
    def excel_file(self, value):
        self.__excel_file = value        

    #property get and set for parquet filter value
    @property
    def parquet_filter(self):
        return self.__parquet_filter
    @parquet_filter.setter
    def parquet_filter(self, value):
        self.__parquet_filter = value

    #property get for worksheets list of values
    @property
    def worksheets(self):
        return self.__worksheets
    @worksheets.setter
    def worksheets(self, value):
        self.__worksheets = value    

    #property get for workbook instance
    @property
    def workbook(self):
        return self.__workbook
    @workbook.setter
    def workbook(self, value):
        self.__workbook = value    

    #property get for parquet data
    @property 
    def parquet_data(self):
        return self.__parquet_data
    @parquet_data.setter
    def parquet_data(self, value):
        self.__parquet_data = value     

    #property get for parquet column
    @property
    def parquet_column(self):
        return self.__parquet_column
    @parquet_column.setter 
    def parquet_column(self, value):
        self.__parquet_column = value   

    #property get and set for default sheet name
    @property
    def default_sheet_name(self):
        return self.__defaulf_sheet_name
    @default_sheet_name.setter
    def default_sheet_name(self, value):
        self.__defaulf_sheet_name = value                           

    #create the workbook instance - PRIVATE
    def create_workbook():
        ParquetWorksheet.workbook = Workbook()

    #read parquet file for data extract - PUBLIC
    def create_worksheets_parquet(parquet_file, excel_file, parquet_filter, default_sheet_name):

        #call the __create_workbook method to create the workbook
        ParquetWorksheet.create_workbook()
        
        #pass the parameters to class varaibles for reuse
        ParquetWorksheet.excel_file = excel_file 
        ParquetWorksheet.default_sheet_name = default_sheet_name

        #read the parquet file to load the data to apply the filter
        parquet_data = pd.read_parquet(parquet_file, engine='fastparquet')

        #call the set_filter_parquet method to set the filter column value
        ParquetWorksheet.set_filter_parquet(parquet_filter, parquet_data)

    #filter the parquet data - PRIVATE
    def set_filter_parquet(parquet_filter, parquet_data):
        
        #filter the parquet on the parquet_data and parquet_filter values
        parquet_column = parquet_data[parquet_filter]

        #call the __apply_parquet_filter to build the worksheets from the parquet filter column
        ParquetWorksheet.apply_parquet_filter(parquet_column)

    #add parquet filter which is the column to filter on - PRIVATE 
    #placeholder for column checking and verification
    def apply_parquet_filter(parquet_column):

        #call the __build_worksheet_list to append the values of the filter to the worksheet list
        ParquetWorksheet.build_worksheet_list(parquet_column)


    #build the list of worksheets to create from the parquet filter column - PRIVATE
    def build_worksheet_list(parquet_column):

        #loop through the parquet column values and build the worksheets list
        worksheets = []
        for column_filter in parquet_column:
            worksheets.append(column_filter)

        #call the __create_worksheets to create the worksheets built from the worksheets list    
        ParquetWorksheet.create_worksheets(ParquetWorksheet.default_sheet_name, worksheets)

    #create the worksheets from the parquet filtered column - PRIVATE
    def create_worksheets(default_sheet_name, worksheets = []):

        ws = ParquetWorksheet.workbook.active.title = default_sheet_name
        
        #loop through the worksheets list and create the worksheets
        for worksheet in worksheets:
            ws = ParquetWorksheet.workbook.create_sheet(worksheet)
            ws.title = worksheet

        #call __save_workbook to save the worksheets and the workbook
        ParquetWorksheet.save_workbook(ParquetWorksheet.excel_file)    

    #save the workbook with the new worksheets created - PRIVATE
    def save_workbook(excel_file):

        #save the workbook
        ParquetWorksheet.workbook.save(excel_file)