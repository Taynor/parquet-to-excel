import json
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font, Fill, Border, Side, Alignment
import pandas as pd
import excel_config

class ExcelCustomStyler:
    def __init__(self, json_specification, json_styling, excel_file, default_sheet_name, 
    parquet_load_file, default_load_sheet, worksheets, content_bold_title_large_font, 
    content_title_large_font, content_bold_title_small_font, content_bold_title_large_alignment_style,
    content_value_medium_font, child_title_small_font, child_bold_title_small_font,
    child_values_small_font, hyperlink_underline_style):
        
        #these fields are available for the user to manipulate
        self.json_specification = json_specification
        self.json_styling = json_styling
        self.excel_file = excel_file
        self.default_sheet_name = default_sheet_name
        self.parquet_load_file = parquet_load_file

        #these fields are hidden and are for PRIVATE use only
        self.worksheets = worksheets
        self.default_load_sheet = default_load_sheet
        self.content_bold_title_large_font = content_bold_title_large_font
        self.content_title_large_font = content_title_large_font
        self.content_bold_title_small_font = content_bold_title_small_font
        self.content_bold_title_large_alignment_style = content_bold_title_large_alignment_style
        self.content_value_medium_font = content_value_medium_font
        self.child_title_small_font = child_title_small_font
        self.child_bold_title_small_font = child_bold_title_small_font
        self.child_values_small_font = child_values_small_font
        self.hyperlink_underline_style = hyperlink_underline_style

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

    #property get and set for excel file value  
    @property
    def excel_file(self):
        return self.__excel_file
    @excel_file.setter
    def excel_file(self, value):
        self.__excel_file = value    

    #property get and set for the default worksheet value
    @property
    def default_sheet_name(self):
        return self.__default_sheet_name
    @default_sheet_name.setter
    def default_sheet_name(self, value):
        self.__default_sheet_name = value 

    #property get and set for parquet file value 
    @property
    def parquet_load_file(self):
        return self.__parquet_file
    @parquet_load_file.setter
    def parquet_file(self, value):
        self.__parquet_file = value 

    #property get for parquet default load sheet value
    @property
    def default_load_sheet(self):
        return self.__default_load_sheet
    @default_load_sheet.setter
    def default_load_sheet(self, value):
        self.__default_load_sheet = value

    #property get and set for worksheets value
    @property 
    def worksheets(self):
        return self.__worksheets
    @worksheets.setter
    def worksheets(self, value):
        self.__worksheets = value

    #property get and set for content bold title large font value
    @property 
    def content_bold_title_large_font(self):
        return self.__content_bold_title_large_font
    @content_bold_title_large_font.setter
    def content_bold_title_large_font(self, value):
        self.__content_bold_title_large_font = value

    #property get and set for content title large font value
    @property 
    def content_title_large_font(self):
        return self.__content_title_large_font
    @content_title_large_font.setter
    def content_title_large_font(self, value):
        self.__content_title_large_font = value

    #property get and set for content bold title small font value
    @property 
    def content_bold_title_small_font(self):
        return self.__content_bold_title_small_font
    @content_bold_title_small_font.setter
    def content_bold_title_small_font(self, value):
        self.__content_bold_title_small_font = value 

    #property get and set for content bold title large alignment style value
    @property 
    def content_bold_title_large_alignment_style(self):
        return self.__content_bold_title_large_alignment_style
    @content_bold_title_large_alignment_style.setter
    def content_bold_title_large_alignment_style(self, value):
        self.__content_bold_title_large_alignment_style = value   

    #property get and set for content value medium font value
    @property 
    def content_value_medium_font(self):
        return self.__content_value_medium_font
    @content_value_medium_font.setter
    def content_value_medium_font(self, value):
        self.__content_value_medium_font = value  

    #property get and set for child title small font value
    @property 
    def child_title_small_font(self):
        return self.__child_title_small_font
    @child_title_small_font.setter
    def child_title_small_font(self, value):
        self.__child_title_small_font = value    

    #property get and set for child bold title small font value
    @property 
    def child_bold_title_small_font(self):
        return self.__child_bold_title_small_font
    @child_bold_title_small_font.setter
    def child_bold_title_small_font(self, value):
        self.__child_bold_title_small_font = value 

    #property get and set for child values small font value
    @property 
    def child_values_small_font(self):
        return self.__child_values_small_font
    @child_values_small_font.setter
    def child_values_small_font(self, value):
        self.__child_values_small_font = value

    #property get and set for hyperlink underline style value
    @property 
    def hyperlink_underline_style(self):
        return self.__hyperlink_underline_style
    @hyperlink_underline_style.setter
    def hyperlink_underline_style(self, value):
        self.__hyperlink_underline_style = value                                                 

    #read in the style sheet to load in the global variables for styling
    def read_style_sheet(json_styling):
        
        #pass the parameters to class variables for reuse
        ExcelCustomStyler.json_styling = json_styling

        #load styles from json_styles
        with open(json_styling) as jst:
            style_config = json.load(jst)

            #load the styles for large content bold title labels
            for content_bold_title_large in style_config["custom_content_bold_title_large"]:
                content_bold_title_large_font = Font(name=content_bold_title_large['font_name'],
                size=content_bold_title_large['font_size'],
                bold=content_bold_title_large['bold'],
                color=content_bold_title_large['font_colour'])

            #load the styles for large content title values
            for content_title_large in style_config["custom_content_title_large"]:   
                content_title_large_font = Font(name=content_title_large['font_name'],
                size=content_title_large['font_size'],
                color=content_title_large['font_colour'])

            #load the style for small content bold title values
            for content_bold_title_small in style_config["custom_content_bold_title_small"]:
                content_bold_title_small_font = Font(name=content_bold_title_small['font_name'],
                size=content_bold_title_small['font_size'],
                bold=content_bold_title_large['bold'],
                color=content_bold_title_small['font_colour'])     

            #load the styles for large content bold title labels alignment
            for content_bold_title_large_alignment in style_config["custom_alignment_content_bold_title_large"]:
                content_bold_title_large_alignment_style = Alignment(horizontal=content_bold_title_large_alignment['horizontal'],
                vertical=content_bold_title_large_alignment['vertical'])   

            #load the styles for medium content values
            for content_value_medium in style_config["custom_content_sheet_font"]:
                content_value_medium_font = Font(name=content_value_medium['font_name'],
                size=content_value_medium['font_size'],
                color=content_value_medium['font_colour'])   

            #load the styles for the small child title values
            for child_title_small in style_config["custom_child_title_small"]:
                child_title_small_font = Font(name=child_title_small['font_name'],
                size=child_title_small['font_size'],
                color=child_title_small['font_colour'])    

            #load the styles for the small child bold title values
            for child_bold_title_small in style_config["custom_child_bold_title_small"]:
                child_bold_title_small_font = Font(name=child_bold_title_small['font_name'],
                size=child_bold_title_small['font_size'],
                bold=child_bold_title_small['bold'],
                color=child_bold_title_small['font_colour'])   

            #load the styles for the small child values
            for child_values_small in style_config["custom_child_values_small"]:
                child_values_small_font = Font(name=child_values_small['font_name'],
                size=child_values_small['font_size'],
                color=child_values_small['font_colour'])    

            #load the styles for the child sheet to content sheet hyperlink
            for hyperlink_underline in style_config["custom_hyperlink_underline"]:
                hyperlink_underline_style = Font(name=hyperlink_underline['font_name'],
                size=hyperlink_underline['font_size'],
                underline=hyperlink_underline['underline'],
                color=hyperlink_underline['font_colour'])
            
            #load the styles for the base sheet borders
            for thin_borders in style_config["custom_thin_borders"]:             
                thin_style = Side(border_style=thin_borders['border_width'],
                color=thin_borders['border_colour'])
                thin_borders_side_style = Border(right=thin_style, left=thin_style)
                thin_borders_top_style = Border(top=thin_style, right=thin_style, left=thin_style)
                thin_borders_bottom_style = Border(bottom=thin_style, right=thin_style, left=thin_style)  
                thin_borders_full_style = Border(top=thin_style, bottom=thin_style, right=thin_style, left=thin_style)
                thin_borders_top_bottom_style = Border(top=thin_style, bottom=thin_style)
                thin_borders_top_bottom_right_style = Border(top=thin_style, bottom=thin_style, right=thin_style)
                thin_borders_top_bottom_left_style = Border(top=thin_style, bottom=thin_style, left=thin_style)
            
            #pass the parameters to class variables for reuse
            ExcelCustomStyler.content_bold_title_large_font = content_bold_title_large_font
            ExcelCustomStyler.content_title_large_font = content_title_large_font
            ExcelCustomStyler.content_bold_title_small_font = content_bold_title_small_font
            ExcelCustomStyler.content_bold_title_large_alignment_style = content_bold_title_large_alignment_style
            ExcelCustomStyler.content_value_medium_font = content_value_medium_font
            ExcelCustomStyler.child_title_small_font = child_title_small_font
            ExcelCustomStyler.child_bold_title_small_font = child_bold_title_small_font
            ExcelCustomStyler.child_values_small_font = child_values_small_font
            ExcelCustomStyler.hyperlink_underline_style = hyperlink_underline_style
    
    #takes the arguments to build the worksheets list in order for the styler
    #to style the worksheets in the dynamic list of worksheets
    def style_worksheets(excel_file, parquet_load_file, json_specification, json_styling):

        #import the styling sheet
        ExcelCustomStyler.read_style_sheet(json_styling)
        
        #pass the parameters to class variables for reuse
        ExcelCustomStyler.excel_file = excel_file
        ExcelCustomStyler.parquet_load_file = parquet_load_file
        ExcelCustomStyler.json_specification = json_specification
        ExcelCustomStyler.json_styling = json_styling

        #read the contents of the parquet loader that has the filter to where the data
        #should be loaded into the respective worksheets
        ExcelCustomStyler.read_parquet_loader(ExcelCustomStyler.parquet_load_file)

    #read parquet file that acts as a pointer to where the data needs to be loaded into
    def read_parquet_loader(parquet_load_file):
        
        #read the content of the parquet loading file
        ExcelCustomStyler.default_load_sheet = pd.read_parquet(parquet_load_file, engine='fastparquet')

        #set the filter column read in from the parquet
        ExcelCustomStyler.create_worksheet_list()
    
    #creates a list of worksheets dynamically created from the parquet filter column
    def create_worksheet_list():
        
        #create an empty to append the column values from the parquet filter
        worksheets = []
         
        #apply the filter to the parquet data loader file to build the list of
        #respective worksheets that will have it's data loaded into
        #there is a bug that means the filter column cannot be passed in as variable
        #it needs to be hardcoded in as an attribute of the dataframe
        ExcelCustomStyler.filter_column = ExcelCustomStyler.default_load_sheet.Question
        for fc in ExcelCustomStyler.filter_column:
            worksheets.append(fc)

        #pass the parameters to class variables for reuse   
        ExcelCustomStyler.worksheets = worksheets

        #load the excel file with the data loaded to start applying styling
        ExcelCustomStyler.load_excel_worksheet(ExcelCustomStyler.excel_file, ExcelCustomStyler.json_specification, ExcelCustomStyler.worksheets)

    #load excel worksheet
    def load_excel_worksheet(excel_file, json_specification, worksheets = []):

        #load the workbook
        excel_template = load_workbook(excel_file)
        
        #Add the content sheet default labels and titles
        for ws in worksheets:
            if ws == worksheets[0]:
                worksheet = excel_template[ws[0:]]
                with open(json_specification) as jsp:
                    spec_config = json.load(jsp)
                    #Add the labels and titles
                    for content_project_details in spec_config["content_project_details"]:
                        worksheet[excel_config.content_sheet_labels[0]] = content_project_details['Project_Label']
                        worksheet[excel_config.content_sheet_labels[1]] = content_project_details['Project_Name']
                        worksheet[excel_config.content_sheet_labels[2]] = content_project_details['Wave_Label']
                        worksheet[excel_config.content_sheet_labels[3]] = content_project_details['Wave_Name']
                        worksheet[excel_config.content_sheet_labels[4]] = content_project_details['Fieldwork_Label']
                        worksheet[excel_config.content_sheet_labels[5]] = content_project_details['Fieldwork_Name']    
                        worksheet[excel_config.content_sheet_labels[6]] = content_project_details['Child_Sheet_Label']
                        worksheet[excel_config.content_sheet_labels[7]] = content_project_details['Question_French_Label']
                        worksheet[excel_config.content_sheet_labels[8]] = content_project_details['Question_English_Label']

                        #Add styling import
                        #Add styling to the large formatting for the content sheet title labels
                        worksheet[excel_config.content_sheet_labels[0]].font = ExcelCustomStyler.content_bold_title_large_font
                        worksheet[excel_config.content_sheet_labels[2]].font = ExcelCustomStyler.content_bold_title_large_font
                        worksheet[excel_config.content_sheet_labels[4]].font = ExcelCustomStyler.content_bold_title_large_font

                        #Add styling to the large formatting for the content sheet title values
                        worksheet[excel_config.content_sheet_labels[1]].font = ExcelCustomStyler.content_title_large_font
                        worksheet[excel_config.content_sheet_labels[3]].font = ExcelCustomStyler.content_title_large_font
                        worksheet[excel_config.content_sheet_labels[5]].font = ExcelCustomStyler.content_title_large_font

                        #Add styling to the small formatting for the content sheet title values
                        worksheet[excel_config.content_sheet_labels[6]].font = ExcelCustomStyler.content_bold_title_small_font
                        worksheet[excel_config.content_sheet_labels[7]].font = ExcelCustomStyler.content_bold_title_small_font
                        worksheet[excel_config.content_sheet_labels[8]].font = ExcelCustomStyler.content_bold_title_small_font
    
        #Add content for child sheets
        excel_template.save(ExcelCustomStyler.excel_file)  