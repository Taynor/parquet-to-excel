#This contains the values to pull in and use for allocating columns, rows and cells
#Pulling these values from the JSON fails even with conversion to string

#cells to load in the child worksheets labels and hyperlinks into the content sheet
child_sheet_hyperlink = ['A6', 'A7', 'A8', 'A9', 'A10', 'A11', 'A12', 'A13', 'A14', 'A15']

#cells to load default content sheet project labels 
content_sheet_labels = ['A1', 'B1', 'A2', 'B2', 'A3', 'B3', 'A5', 'B5', 'C5', 'A6']

#columns to format for the content sheet
content_sheet_labels_dimensions = ['A', 'B', 'C']

#content sheet values cells
content_sheet_values = ['B6', 'B7', 'B8', 'B9', 'B10', 'B11', 'B12', 'B13', 'B14',
'C6', 'C7', 'C8', 'C9', 'C10', 'C11', 'C12', 'C13', 'C14']

#cells to load regions in child sheets
child_sheet_regions = ['C9', 'D9', 'E9', 'F9', 'G9', 'H9', 'I9', 'J9', 'K9', 'L9', 'M9', 'N9', 'O9'
, 'B10', 'C8', 'D8', 'E8', 'F8', 'G8', 'H8', 'I8', 'J8', 'K8', 'L8', 'M8', 'N8', 'O8']

#cells to load value data to in the child sheets
child_sheet_data_values_columns = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']
child_sheet_data_values_rows = [11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23]

##cell to load the back to content sheet hyperlink in the child sheet
content_sheet_hyperlink = 'A9'

#cells to load the child sheet default project labels
child_sheet_labels = ['I1', 'C2', 'H2', 'B10', 'B10:B11', 'D10', 'C', 'D', 'E', 'F', 'G', 'H',
'I', 'J', 'K', 'L', 'M', 'N', 'O', 'A9', 'B', 9]

#child sheet formatting cells for adding the individual questions to the each worksheet
child_sheet_question_labels = ['B3', 'H3', 'B4', 'H4']

#child sheet question labels clean up parquet write bug, this is required before any formatting is done after data load
child_sheet_question_labels_cleanup = ['B4', 'B5', 'B6', 'B7', 'B8', 'B9', 'H4', 'H5', 'H6', 'H7', 'H8', 'H9', 'H10', 'H11']

#child sheet answer opyions to the questions on each sheet
child_sheet_answer_labels = ['B12', 'B13', 'B14', 'B15', 'B16', 'B17', 'B18', 'B19', 'B20',
'B21', 'B22', 'B23', 'B24', 'B25', 'B26', 'B27', 'B28', 'B29', 'B30', 'B31', 'B32', 'B33']

#child sheet cells for the custom base details per sheet
child_sheet_base_custom_details = ['B5', 'H5']

#cells to add for the base sheet
base_sheet_labels = ['B12', 'B13', 'B14', 'B15', 'B16', 'B17', 'B18', 'B19', 'B20', 'B21',
'B22', 'B23', 'B24', 'B25', 'B26', 'B27', 'B28', 'B29', 'B30', 'B31', 'B32', 'B33', 'B34',
'B35', 'B36', 'B37', 'B38', 'B39', 'B40', 'B41', 'B42', 'B43', 'B44', 'B45', 'B46', 'B47',
'B48', 'B49', 'B50', 'B51', 'B52', 'B53', 'B54', 'B55', 'B56', 'B57', 'B58', 'B59', 'B60',
'B61', 'B62', 'B63', 'B64', 'B65', 'B66', 'B67', 'B68', 'B69', 'B70', 'B71', 'B72', 'B73',
'B74', 'B75', 'B76', 'B77', 'B78', 'B79', 'B80', 'B81', 'B82', 'B83', 'B84', 'B85', 'B86',
'B87', 'B88', 'B89', 'B90', 'B91', 'B92', 'B93', 'B94', 'B95']

#the cell range for adding the base sheet border above the regions columns
base_sheet_borders = ['B12:B95']

#the cell range for the base child label and detail for the merge of cells
child_sheet_base_custom_merge_cells = ['B5:F5', 'H5:L5']

#the cell range for adding borders around the Total values in the child sheets
child_sheet_total_borders = ['C10', 'D10', 'E10', 'F10', 'G10', 'H10', 'I10', 'J10', 'K10', 'L10',
'M10', 'N10', 'O10', 'C11', 'D11', 'E11', 'F11', 'G11', 'H11', 'I11', 'J11', 'K11', 'L11',
'M11', 'N11', 'O11']

#merge cells for the questions content loaded into each worksheet. This will customise
#the alignment for the dynamic content. The first is for French and the second for English
child_sheet_question_content = ['B3:F3', 'H3:L3']

#the parquet location for the content sheet data adapter parquet file
content_sheet_parquet = 'C:\python-excel-template-git\python-excel-template\parquets\content_sheet\content_sheet.snappy.parquet'
base_sheet_parquet = 'C:\python-excel-template-git\python-excel-template\parquets\sheet_base\sheet_base.snappy.parquet'
qb1_sheet_parquet = 'C:\python-excel-template-git\python-excel-template\parquets\qb1_sheet\qb1_sheet.snappy.parquet'
qb1r_sheet_parquet = 'C:\python-excel-template-git\python-excel-template\parquets\qb1r_sheet\qb1r_sheet.snappy.parquet'
qb2_sheet_parquet = 'C:\python-excel-template-git\python-excel-template\parquets\qb2_sheet\qb2_sheet.snappy.parquet'
qb2r_sheet_parquet = 'C:\python-excel-template-git\python-excel-template\parquets\qb2r_sheet\qb2r_sheet.snappy.parquet'
qb3_sheet_parquet = 'C:\python-excel-template-git\python-excel-template\parquets\qb3_sheet\qb3_sheet.snappy.parquet'
qb3r_sheet_parquet = 'C:\python-excel-template-git\python-excel-template\parquets\qb3r_sheet\qb3r_sheet.snappy.parquet'
qb4_sheet_parquet = 'C:\python-excel-template-git\python-excel-template\parquets\qb4_sheet\qb4_sheet.snappy.parquet'
qb5_sheet_parquet = 'C:\python-excel-template-git\python-excel-template\parquets\qb5_sheet\qb5_sheet.snappy.parquet'

#the excel template configuration
excel_template_configuration = 'C:\python-excel-template-git\python-excel-template\excel_template_configuration.xlsx'

#the excel template
excel_template = 'C:\python-excel-template-git\python-excel-template\excel_template.xlsx'

#the JSON config file to apply custom content to the worksheets
json_specification = 'excel_specification.json'

#the JSON styles file to add custom styling to the worksheets
json_styling = 'excel_styles.json'