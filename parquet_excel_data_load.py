from openpyxl import Workbook
from openpyxl import load_workbook
import pandas as pd
import os
import pathlib
import glob

class ParquetExcelDataLoad:
    def __init__(self, parquet_path, parquet_file, excel_file, parquet_list, parquet_subdirectories,
    worksheets, parquet_file_pattern, parquet_folders):
        
        #these fields are available for the user to manipulate
        self.parquet_path = parquet_path
        self.parquet_file = parquet_file
        self.excel_file = excel_file

        #these fields are hidden and are for PRIVATE use only
        self.parquet_list = parquet_list
        self.parquet_subdirectories = parquet_subdirectories
        self.worksheets = worksheets
        self.parquet_file_pattern = parquet_file_pattern
        self.parquet_folders = parquet_folders
    
    #property get and set for parquet path value
    @property
    def parquet_path(self):
        return self.__parquet_path
    @parquet_path.setter
    def parent_path(self, value):
        self.__parquet_path = value

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


            