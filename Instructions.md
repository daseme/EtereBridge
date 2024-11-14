# EtereBridge: Traffic Data Transformation Utility

**EtereBridge** is a Python-based utility for transforming ad traffic data from the Etere system into a format compatible with your in-house tracking and reporting system. This guide provides step-by-step instructions for setting up and using EtereBridge.

---

## Directory Setup

Before using the script, set up the following directory structure:

- **templates**: This folder contains the template Excel file (used as a base for each transformation).
- **input/monthly_files**: Place the monthly input CSV files from Etere here.
- **output**: This is where EtereBridge will save the processed output files.

For example:

project-folder/ │ ├── templates/ │ └── Template_File_B.xlsx │ ├── input/ │ └── monthly_files/ │ ├── January_2024.csv │ └── February_2024.csv │ └── output/



### File Naming Conventions

- **Template File (Template_File_B.xlsx)**: Place this template in the `templates` directory. It will be reused each month to ensure consistent formatting and structure.
- **Monthly Input Files (CSV)**: Each month, place the new CSV input files from Etere in `input/monthly_files/`.
- **Output Files**: EtereBridge saves processed files in the `output` directory with names based on the input file name and the current date (e.g., `January_2024_Processed_2024-01-01.xlsx`).

---

## Steps to Use EtereBridge

### Step 1: Install Python and Required Libraries

1. Make sure you have Python installed on your computer. You can download it from [python.org](https://www.python.org/downloads/).
2. You need two Python libraries: `pandas` and `openpyxl`. Open your command prompt (Windows) or terminal (Mac/Linux) and run:

    ```bash
    pip install pandas openpyxl
    ```

---

### Step 2: Prepare Your Files

1. Place the template file, `Template_File_B.xlsx`, in the `templates` directory.
2. Place the monthly CSV input files from Etere in the `input/monthly_files` directory.

---

### Step 3: Run the EtereBridge Script

1. Open a command prompt or terminal.
2. Navigate to the folder where the EtereBridge script is saved. For example:

    ```bash
    cd path/to/your/project-folder
    ```

3. Run the script by typing:

    ```bash
    python EtereBridge.py
    ```

---

### Step 4: Choose Processing Mode and Answer Prompts

- **Processing Mode**: The script will prompt you to choose:
    - **Process All Files**: Select this option to process every CSV file in the `input/monthly_files` directory.
    - **Select Files Individually**: Choose this option to pick files one at a time.

- After selecting the processing mode, EtereBridge will ask for additional information to fill specific columns in the output:
    - **Enter Billing Type (Calendar or Broadcast):** Type "Calendar" or "Broadcast" based on your needs.
    - **Enter Revenue Type:** Type in a specific revenue type (e.g., "Ad Sales" or "Subscription").
    - **Is this an agency order? (Yes or No):** Type "Yes" or "No" depending on whether this is an agency order.

These answers will automatically populate the corresponding columns in each processed output file.

---

### Step 5: Check the Output Files

When the script finishes, it will save each processed file as a new Excel file in the `output` directory. The filename will include the original input file name and the current date, for example: output/January_2024_Processed_2024-01-01.xlsx


Open each output file in Excel to review the data and ensure the formatting and transformations were applied correctly.

---

## Important Notes

- **Do Not Close the Command Prompt or Terminal** while the script is running, as you need it open to answer the prompts.
- **Error Handling**: If you encounter error messages, ensure that the file paths are correct and that you’ve installed the required libraries (`pandas` and `openpyxl`).
- **Organize Files**: Make sure to keep the template and input files in their respective directories for consistent processing each month.
- **Re-run Options**: If you want to re-process files or modify your answers, simply re-run the script.

---

## Summary

1. Set up the directory structure with folders for templates, input files, and output files.
2. Place the template file in `templates/`, and monthly input files in `input/monthly_files`.
3. Run EtereBridge in the command prompt or terminal.
4. Choose to process all files or select files individually, then answer the prompts.
5. Open each processed file in `output/` to review the results.

EtereBridge is designed to streamline and standardize your monthly data transformations, providing consistent, high-quality outputs for tracking and reporting.


