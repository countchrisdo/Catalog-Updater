# Catalog Updater - The CLI Data Copy and Merge Tool

![Python Version](https://img.shields.io/badge/python-3.6%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)

## Overview

The CLI Data Copy and Merge Tool is a command-line interface (CLI) program written in Python. Initially designed to copy millions of products from a distrubuter's **Catalog** to a Seller's **Cost Price Files** and **Resale Price Files**. This tool allows you to copy data from specific columns in multiple different files and paste them into one large template file. This was created for my own specific use at work but it's published for anyone to use because this code can be edited fairy easily to do a lot of different Copy-Pasting tasks in Microsoft Excel. It's especially useful when you have structured data in multiple source files that need to be consolidated into a single template. 

## Features

- Copy data from specific columns in multiple files.
- Merge the copied data into a single template file.
- Customize the template file and column mappings.
- Easily automate data consolidation tasks.

## Installation

Before you can use the CLI Data Copy and Merge Tool, you need to install Python on your system if it's not already installed. You can download Python from the [official website](https://www.python.org/downloads/).

1. Clone this GitHub repository to your local machine or download the source code as a ZIP file.

```bash
git clone https://github.com/countchrisdo/Catalog-Updater
```

2. Navigate to the project directory:

```bash
cd Catalog-Updater
```

3. Install the required dependencies using pip:

```bash
pip install -r requirements.txt
```

## Usage

To use the CLI Data Copy and Merge Tool, follow these steps:

1. Prepare your source data files in the **Inbox** folder, and ensure that they have a consistent format (e.g., CSV, Excel).

2. Create a template file with placeholders where you want to insert the copied data. Name the file "Template" or change the name in main.py line 10.

3. You can enter the 5 letters of your desired columns inside the *dialog*'s function **lst[]** 

5. Run the CLI tool, review the terminal output to ensure the files were imported correctly, and then press **Enter** to continue.

6. The tool will copy data from the specified columns in your source files, merge it into the template, and save the result in the **Output** folder.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Special thanks to Openpyxl who have made this tool possible with their incredible Excel + Python integration

## Contact

If you have any questions or suggestions, feel free to contact me at [crburwell@yahoo.com](mailto:crburwell@yahoo.com).

Happy data merging!

# Todo

- ~~Before running, display imported files~~
- ~~Add Color to print statements for readability in Terminal~~
- Let user define file paths when running the program
- Let user define column mappings in the terminal
