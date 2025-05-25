# Excel Replicator using Python
![Sunset](https://github.com/user-attachments/assets/3b26020f-d964-4e33-b1b9-38b7de34ce1b)
## Introduction

**Excel Replicator using Python** is a robust and user-friendly application designed to streamline the process of copying and consolidating data across multiple Excel sheets and workbooks. Built with Python and Tkinter, this tool simplifies the data transfer process, making it efficient and error-free.

## Problems Solved

Handling large Excel files with multiple sheets can be cumbersome and error-prone. Manually copying data between sheets or workbooks is time-consuming and can lead to inconsistencies. This application addresses these issues by automating the replication process, ensuring data accuracy and saving valuable time.

## Uses and Use Cases

- **Data Consolidation**: Merge data from various Excel sheets into a single workbook.
- **Data Backup**: Create backups of important Excel sheets to prevent data loss.
- **Data Migration**: Transfer data between different Excel files seamlessly.
- **Data Analysis Preparation**: Prepare data from multiple sources for analysis by consolidating it into one sheet or workbook.

## Features

- **Entire Workbook Replication**: Copy all sheets from a source workbook to a target workbook.
- **Sheet-Specific Replication**: Copy specific sheets from a source workbook to a target workbook.
- **User-Friendly Interface**: Simple and intuitive GUI built with Tkinter.
- **Error Handling**: Comprehensive error handling to ensure smooth operation.
- **View Data**: Preview data from selected sheets before replication.

## How to Set Up

### Prerequisites

Ensure you have the following Python libraries installed:
- tkinter
- openpyxl
- pandas
- Pillow

You can install these libraries using pip:

```sh
pip install tkinter openpyxl pandas Pillow
```

### Code Structure

The project is divided into several modules for better organization and maintainability:

1. **main.py**: The entry point of the application.
2. **excel_replicator.py**: Contains the core logic for Excel sheet replication.
3. **file_browsing.py**: Handles file selection dialogs and loading sheet names.
4. **viewing.py**: Manages the functionality to preview Excel sheets.
5. **ui_components.py**: Contains the code for the user interface elements.
6. **constants.py**: Defines constant values used across the application.
7. **assets**: Directory for storing images and icons used in the application.

### Key Code Structures

- **ExcelReplicatorApp**: Main class that initializes the GUI and handles user interactions.
- **browse_source_file()**: Opens a file dialog to select the source Excel file.
- **browse_target_file()**: Opens a file dialog to select the target Excel file.
- **load_sheet_names()**: Loads the names of sheets in a selected Excel file.
- **view_file()**: Previews the content of a selected sheet.
- **pull_data()**: Copies data from the source sheet to the target sheet.
- **replicate_data()**: Replicates data based on user selections, including the option to copy entire workbooks.

## How to Use

1. **Run the Application**:
   Execute the `main.py` file to start the application.

2. **Select Source and Target Files**:
   Use the 'Browse' buttons to select the source and target Excel files.

3. **Choose Sheets or Entire Workbook**:
   Select the specific sheets to replicate or choose the 'Entire File' option to replicate the whole workbook.

4. **Replicate Data**:
   Click the 'Pull' or 'Push' buttons to perform the data replication.

## Crafted with Love

Crafted with Love by Sam Naveenkumar .V ❤️

© 2024 All rights reserved.

---
