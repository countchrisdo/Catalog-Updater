"""Catalog Updater - config.py""" 

# Directory containing input Excel files
INPUT_DIR = "Inbox"

# Name of Template file to search for
template_files = ["Template.xlsx", "Template.xlsm"]

# Output path for merged results
OUTPUT_PATH = "Outbox/Output.xlsx"

# Columns to extract from source files and their target columns in the template
# Format: Each item in the list represents a column in the template. Input a 0 for a column you want to skip and input the letter of the column you want to extract from the source file. The first index of the list should always be 0.
COLUMN_MAPPING = ["0", "A", "0", "O", "0", "P", "Q"]

# TODO Chunk size
# CHUNKSIZE = 5000
