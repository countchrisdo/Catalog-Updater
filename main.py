import os
import sys
from tkinter import N
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import Font

#TODO Move a bunch of these functions to another file to clean up main.py
# Importing Spreadsheets
wb_template = load_workbook('Dummy Template.xlsx')
#wb_data = load_workbook('Inbox/Dummy_FILE_A.xlsx')

directory = 'Inbox'

inbox_files = []

#Note: You can iterate through a list with forloop automatically

for filename in os.listdir(directory):
    f = os.path.join(directory, filename)
    # checking if it is a file
    if "~" in f:
        print("Warning: File is open")
    elif os.path.isfile(f):
        print(f)
        inbox_files.append(f)
    else:
        print("incorrect file found, or no files found. try again")

wb_data = load_workbook(inbox_files[0])

def start():
    txt = input("Does this look correct?: (Y/N)")
    if txt == "Y" or "y":
        return
    else:
        sys.exit()
# start()

def main():
    global wb_data

    import_num = 0
    for filename in os.listdir(directory):
        f = os.path.join(directory, filename)
        # checking if it is a file
        if "~" in f:
            print("Warning: File is open")
        elif os.path.isfile(f):
            import_num += 1
    print(f"You have: {import_num} files available")

    ws_template = wb_template.active
    ws_data = wb_data.active

    # Copying sheets to match import number
    for file_index in range(1, import_num):
        new_sheet = wb_template.copy_worksheet(ws_template)
        new_sheet.title = f"Sheet{file_index + 1}"
        print(f"Initializing {new_sheet.title}")

    row_count = len(tuple(ws_data.rows))

    # Asking user for the columns to use
    col_list = [0]
    def dialog(lst):
        # Maybe change to list what each input is for / Cost / Id / ect
        # print("Input the Letter of the Columns you want scraped")
        # print("Please enter one at a time:")
        # print("Enter a zero to leave it blank:")
        # for x in range(0,6):
        #     txt = input("-")
        #     lst.append(txt)
        #     print(txt)
        # dummy info for the lift
        lst = ["0", "A", "0", "O", "0", "P", "Q"]
        print(f"Selected Columns:{lst}")
        return (lst)
    col_list = dialog(col_list)

    def copypaster(x, y):
        print("Copypaster function run")
        if x == "0":
            print("Leaving Field Blank")

        else:
            # for cell in data sheet's column X, print data, copy data, print again
            # y += 1
            for cell in range(2, row_count + 1):
                data_cell = ws_data[x + str(cell)]
                print(data_cell.value)
                # +4 on the position of the new cell to get under the headers
                new_cell = ws_template[get_column_letter(y) + str(cell + 4)]
                new_cell.value = data_cell.value
                print(new_cell.value)

    def old_colswitcher(num, str):
        # for each column in data sheet, run copypaster then switch column and run again
        for column in range(0, 5):
            # increment character A, B, C... etc
            num += 1
            str = get_column_letter(num)

            copypaster(str)

    # Take in list and switch to each column as needed to read information
    def colswitcher(lst):
        print("Column Switcher function run")
        # Switches 6 times exactly starting at Number "1" AKA "A"
        # should I make number of switches vary on array length?
        for column in range(1, 7):
            # copypaster takes in a string value at x, number at y
            print(f"Switched to index number {column} - {lst[column]}")
            copypaster(lst[column], column)
    colswitcher(col_list)

    # switch to next sheet and file

    #for each file in the inbox_files list
    #-1 from length because the ocde has already run once by default
    idx = 0
    for file in range(0,len(inbox_files)-1):
        idx += 1
        print(idx)
        ws_template = wb_template["Sheet" + str(idx + 1)]
        print(f"Switched to {ws_template}")
        wb_data = inbox_files[idx]
        print(f"Switched to file {inbox_files[idx]}")
        colswitcher(col_list)

    # Saving the filled in Template as a new file

    def end():
        txt = input("Would you like to Save your Progress?: (Y/N)")
        if txt == "Y" or "y":
            return
        else:
            sys.exit()
    # end()

    # put this code inside the end() function
    wb_template.save('Outbox/Output.xlsx')
    print("File exported in /Outbox")


if __name__ == "__main__":
    main()