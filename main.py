import os
import sys
from termcolor import colored, cprint
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
import time

# TODO: Move a bunch of these functions to another file to clean up main.py
print("Program Starting")
begin = time.time()

# Check for "Template.xlsx" and "Template.xlsm" files
template_files = ["Template.xlsx", "Template.xlsm"]
found_template = None
for filename in template_files:
    if os.path.isfile(filename):
        found_template = filename
        print(f"Found template file: {found_template}")
        break

# If a template file is found, load it
if found_template:
    print(f"Loading {found_template} into memory.")
    wb_template = load_workbook(found_template)
else:
    print("No template file found.")

#Move from Root to Inbox folder
directory = 'Inbox'
inbox_files = []
for filename in os.listdir(directory):
    f = os.path.join(directory, filename)
    # checking if it is a xls (xlsx, xlsm) file
    if "xls" in f:
        inbox_files.append(f)
if inbox_files:
    # Load the first file in the inbox_files list at index [0]
    wb_data = load_workbook(inbox_files[0])
else:
    print("No valid files found in the directory.")

def main():
    global wb_data

    print("Catalog Updater Running")
    print(colored("You have:", "cyan"), colored(f"{len(inbox_files)} files available"))
    print(inbox_files)

    ws_template = wb_template.active
    ws_data = wb_data.active
    row_count = len(tuple(ws_data.rows))

    # Copying sheets to match import number
    for file_index in range(1, len(inbox_files)):
        new_sheet = wb_template.copy_worksheet(ws_template)
        new_sheet.title = f"Sheet{file_index + 1}"
        print(f"Initializing {new_sheet.title}")

    col_list = [0]
    def dialog(lst):
        # Idea: change to list what each input is for / Cost / Id / ect

        # This commented out section is for user input of columns
        # print("Input the Letter of the Columns you want scraped")
        # print("Please enter one at a time:")
        # print("Enter a zero to leave it blank:")
        # for x in range(0,6):
        #     txt = input("-")
        #     lst.append(txt)
        #     print(txt)
        
        # Hardcoding the list of Columns in the code for ease of use right now
        lst = ["0", "A", "0", "O", "0", "P", "Q"]
        print(f"Selected Columns:{lst}")
        return (lst)
    col_list = dialog(col_list)

    def copypaster(x, y):
        print("Copypaster run")
        if x == "0":
            print("Leaving Column Blank")
        else:
            # for each cell in data sheet's column X, copy data and paste it in template cell
            for cell in range(2, row_count + 1):
                data_cell = ws_data[x + str(cell)]
                # +4 on the position of the new cell to get under the headers
                new_cell = ws_template[get_column_letter(y) + str(cell + 4)]
                new_cell.value = data_cell.value

    # Take in list and switch to each column as needed to read information
    def colswitcher(lst):
        print("Column Switcher run")
        # Switches 6 times exactly starting at Number "1" AKA "A"
        # IDEA: make number of switches vary on array length?
        for column in range(1, 7):
            print(f"Switched to index number {column} - {lst[column]}")
            # copypaster takes in a string value at x, number at y
            copypaster(lst[column], column)
    colswitcher(col_list)

    # switching to next sheet and file
    # -1 from length because colswitcher has already run once by default
    idx = 0
    for file in range(0, len(inbox_files)-1):
        idx += 1
        ws_template = wb_template["Sheet" + str(idx + 1)]
        print(f"Writing to {ws_template}")
        wb_data = load_workbook(inbox_files[idx])
        ws_data = wb_data.active
        print(f"Scanning file {inbox_files[idx]}")
        colswitcher(col_list)

    # Saving the filled in Template as a new file
    #End function is not called currently
    def end():
        txt = input("Would you like to Save your Progress?: (Y/N)")
        if txt == "Y" or "y":
            return
        else:
            sys.exit()
    # end()

    # TODO put this code inside the end() function
    wb_template.save('Outbox/Output.xlsx')
    print(colored("File exported in:", "cyan"), colored("/Outbox", "white"))

    end = time.time()
    print(colored("Total runtime:", "cyan"),
        colored(f"{end - begin} seconds", "white"))


if __name__ == "__main__":
    main()
