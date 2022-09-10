import os
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

# Importing Spreadsheets
wb_template = load_workbook('Dummy Template.xlsx')
wb_data = load_workbook('Inbox/Dummy_File_A.xlsx')
directory = 'Inbox'


def main():
    # iterate over files in named directory
    print("Checking Inbox Folder")
    import_num = 0
    for filename in os.listdir(directory):

        f = os.path.join(directory, filename)
        # checking if it is a file
        if "~" in f:
            print("Warning: File is open")
        elif os.path.isfile(f):
            import_num += 1
            print(f)
    print(f"You have imported: {import_num} files")

    # Setting to the first template sheet
    ws_template = wb_template.active
    ws_data = wb_data.active
    print("Initializing the Sheet1")

    # Copying sheet1 to import number
    for file_index in range(1, import_num + 1):
        print("Copying Worksheet template")
        new_sheet = wb_template.copy_worksheet(ws_template)
        new_sheet.title = f"Sheet{file_index + 1}"
        print(f"Initializing {new_sheet.title}")

    # Insert data to sheet1
    # TODO convert to a for loop or something
    ws_template['A6'].value = ws_data['A2'].value
    row1 = ws_data['A2:A10']
    print(row1)

    # switch to sheet2
    ws_template = wb_template["Sheet2"]
    # import data from catalog B?
    # insert data to sheet2
    ws_template['A6'].value = "Test Value"
    ws_template['A7'].value = "On Sheet2"

    # Saving the filled in Template as a new file
    wb_template.save('Outbox/Output.xlsx')
    print("File exported in /Outbox")


if __name__ == "__main__":
    main()
