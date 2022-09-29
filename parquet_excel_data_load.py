from openpyxl import Workbook
from openpyxl import load_workbook
import pandas as pd
import os
import pathlib
import glob

class ParquetExcelDataLoad:
    def __init__(self, parquet_path, parquet_load_file, excel_file, parquet_list, parquet_subdirectories,
    worksheets, parquet_file_pattern, parquet_folders, parquet_filter, default_load_sheet, filter_column):
        
        #these fields are available for the user to manipulate
        self.parquet_path = parquet_path
        self.parquet_file = parquet_file
        self.excel_file = excel_file
        self.parquet_filter = parquet_filter

        #these fields are hidden and are for PRIVATE use only
        self.parquet_list = parquet_list
        self.parquet_subdirectories = parquet_subdirectories
        self.worksheets = worksheets
        self.parquet_file_pattern = parquet_file_pattern
        self.parquet_folders = parquet_folders
        self.default_load_sheet
        self.filter_column
    
    #property get and set for parquet path value
    @property
    def parquet_path(self):
        return self.__parquet_path
    @parquet_path.setter
    def parent_path(self, value):
        self.__parquet_path = value

    #property get and set for parquet file value 
    @property
    def parquet_load_file(self):
        return self.__parquet_file
    @parquet_load_file.setter
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

    #property get and set for parquet list value
    @property
    def parquet_list(self):
        return self.__parquet_list
    @parquet_list.setter 
    def parquet_list(self, value):
        self.__parquet_list = value

    #property get and set for parquet subdirectories value
    @property
    def parquet_subdirectories(self):
        return self.__parquet_subdirectories
    @parquet_subdirectories.setter
    def parquet_subdirectories(self, value):
        self.__parquet_subdirectories = value

    #property get and set for worksheets value
    @property 
    def worksheets(self):
        return self.__worksheets
    @worksheets.setter
    def worksheets(self, value):
        self.__worksheets = value
    
    #property get and set for parquet file pattern value
    @property
    def parquet_file_pattern(self):
        return self.parquet_file_pattern
    @parquet_file_pattern.setter 
    def parquet_file_pattern(self, value):
        self.parquet_file_pattern = value  

    #property get and set for parquet folder value     
    @property
    def parquet_folder(self):
        return self.__parquet_folder
    @parquet_folder.setter 
    def parquet_folder(self, value):
        self.__parquet_folder = value

    #property get for parquet default load sheet value
    @property
    def default_load_sheet(self):
        return self.__default_load_sheet
    @default_load_sheet.setter
    def default_load_sheet(self, value):
        self.__default_load_sheet = value     

    #property get for filter column value
    @property
    def filter_column(self):
        return self.__filter_column
    @filter_column.setter
    def filter_column(self, value):
        self.__filter_column = value           

    #the main function that takes the arguments that will perform the data load
    #by taking the pararmeter values and using them as arguments for the downstream functions
    def load_parquet_data(parquet_load_file, excel_file, parquet_path, parquet_filter, parquet_subdirectories=[], parquet_folders=[], parquet_list=[], parquet_file_pattern = '\\*.parquet'):
        
        #pass the parameters to class variables for reuse
        ParquetExcelDataLoad.parquet_load_file = parquet_load_file
        ParquetExcelDataLoad.parquet_filter = parquet_filter

        #create the list of parquet subdirectories
        for files in os.listdir(parquet_path):
            parquet_subdirectories.append(files)     

        #create the parquet list
        for parquet_directory in parquet_subdirectories:

            #the parquet subfolder to a string compare on and add to the list
            parquet_folder = parquet_path + parquet_directory
            parquet_folders.append(parquet_folder)

            #search through the sub folder and find the parquets
            parquets = glob.glob(parquet_folder + parquet_file_pattern)

            #append the list of parquets to the parquet list
            for p in parquets:
                parquet_list.append(p)  

        #read the contents of the parquet loader that has the filter to where the data
        #should be loaded into the respective worksheets
        ParquetExcelDataLoad.read_parquet_loader(ParquetExcelDataLoad.parquet_load_file)                      

    #read parquet file that acts as a pointer to where the data needs to be loaded into
    def read_parquet_loader(parquet_load_file):
        
        #read the content of the parquet loading file
        ParquetExcelDataLoad.default_load_sheet = pd.read_parquet(parquet_load_file, engine='fastparquet')

        #set the filter column read in from the parquet
        ParquetExcelDataLoad.set_filter_parquet(ParquetExcelDataLoad.parquet_filter, ParquetExcelDataLoad.worksheets)

    #sets the filter and loads values of the worksheets that will have its data loaded into
    def set_filter_parquet(parquet_filter, worksheets=[]):
        
        #apply the filter to the parquet data loader file to build the list of
        #respective worksheets that will have it's data loaded into
        #there is a bug that means the filter column cannot be passed in as variable
        #it needs to be hardcoded in as an attribute of the dataframe
        ParquetExcelDataLoad.filter_column = ParquetExcelDataLoad.default_load_sheet.Question
        for fc in ParquetExcelDataLoad.filter_column:
            pass

    #loads the data from the parquet filter and parquet loader file into the respective
    #worksheets using the arguments from the upstream functions
    def load_parquet_content(excel_file, parquet_list, worksheets=[]):
        pass

    #save the workbook with the new worksheets created - PRIVATE
    #def save_workbook(excel_file):

        load_workbook(excel_file)

        #save the content to the Excel workbook file
        #excel_writer.save()
        #excel_writer.close()
            