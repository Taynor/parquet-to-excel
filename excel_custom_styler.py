import json
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font, Fill, Border, Side, Alignment

class ExcelCustomStyler:
    def __init__(self, json_specification, json_styling):
        
        #these fields are available for the user to manipulate
        self.json_specification = json_specification
        self.json_styling = json_styling

    #property get and set for json specification path value  
    @property
    def json_specification(self):
        return self.__json_specification
    @json_specification.setter
    def json_specification(self, value):
        self.__json_specification = value
    
    #property get and set for json styling path value
    @property
    def json_styling(self):
        return self.__json_styling
    @json_styling.setter
    def json_styling(self, value):
        self.__json_sytyling = value

    def read_style_sheet():
        pass
