### Code to access directory
import os
# assign directory
directory = 'Inbox'
 
# iterate over files in named directory
print("Checking Inbox Folder")
for filename in os.listdir(directory):
    f = os.path.join(directory, filename)
    # checking if it is a file
    if "~" in f:
        print("Warning: File is open")
    elif os.path.isfile(f):
        print(f)