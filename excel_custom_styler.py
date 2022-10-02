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
