from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
print('main.py Running')

# Import the Template and Catalog A here
wb_template = load_workbook('Dummy Template.xlsx')
wb_data = load_workbook('Inbox/Dummy_File_A.xlsx')

# Number of data files being imported
import_num = 1

### Setting to the first template sheet 
ws_template = wb_template.active
ws_data = wb_data.active
print("Initializing the first sheet: ")
print(ws_template)

### creating new sheets inside template
## convert into a for loop based on the import number at some point?
new_sheet = wb_template.copy_worksheet(ws_template)
new_sheet.title = "Sheet2"
print("Initializing the 2nd sheet:")
print(wb_template['Sheet2'])

### insert data to sheet1
## convert to a for loop
ws_template['A6'].value = ws_data['A2'].value
ws_template['A7'].value = ws_data['A3'].value

#switch to sheet2
ws_template = wb_template["Sheet2"]
#import data from catalog B

### insert data to sheet2
ws_template['A6'].value = "Test Value"
ws_template['A7'].value = "On Sheet2"



### Saving the filled in Template as a new file
wb_template.save('Outbox/Output.xlsx')
print("File exported in /Outbox")
print(wb_template.sheetnames)