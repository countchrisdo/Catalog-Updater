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
    print("Checking Inbox Folder...")
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

    #inputs
    input("Does this look correct?: (Y/N)")

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
    print("Writing to sheet1...")
    row_count = len(tuple(ws_data.rows))

    def copypaster(x):
        # for cell in column X, print data, copy data, print again
        for cell in range(2, row_count + 1):
            data_cell = ws_data[x + str(cell)]
            print(data_cell.value)
            # +4 on the position of the new cell to get under the headers
            new_cell = ws_template[x + str(cell + 4)]
            new_cell.value = data_cell.value
            print(new_cell.value)
    # for each column in data sheet, run copypaster then switch column and run again
    def colswitcher(num, str):
        for column in range(0,2):
            #increment character A, B, C... etc
            num += 1
            str = get_column_letter(num)
            
            if 1 == 1:
                copypaster(str)
            else:
                colswitcher(str)
   
    # switch to next sheet2 and file
    def fileswitcher():
        sheet_num = 1
        for files in range(0,1):
            ws_template = wb_template["Sheet" + str(sheet_num)]
 
    #starting places inputed into functions
    copypaster("A")
    colswitcher(1, "A")
    # fileswitcher()

    # Saving the filled in Template as a new file
    wb_template.save('Outbox/Output.xlsx')
    print("File exported in /Outbox")


if __name__ == "__main__":
    main()
