import os
import glob
import csv

# Mapping of language keywords to language codes
LANGUAGE_MAPPING = {
    "Chinese": "M",
    "Filipino": "T",
    "Hmong": "Hm",
    "South Asian": "SA",
    "Vietnamese": "V",
    # Add more mappings as needed
}

# Default language if no keyword is found
DEFAULT_LANGUAGE = "E"  # English

def extract_language_from_rowdescription(rowdescription):
    """
    Extract the language code from the rowdescription using keyword matching.
    """
    for keyword, code in LANGUAGE_MAPPING.items():
        if keyword in rowdescription:
            return code
    return DEFAULT_LANGUAGE

def process_file(file_path):
    """
    Process a single file to extract languages from the rowdescription column.
    """
    with open(file_path, mode='r', encoding='utf-8-sig') as file:  # Use utf-8-sig to handle BOM
        reader = csv.DictReader(file)
        for row in reader:
            # Extract rowdescription from the None key (last element in the list)
            if None in row and isinstance(row[None], list) and len(row[None]) > 3:
                rowdescription = row[None][3]  # rowdescription is the 4th element in the list
                language = extract_language_from_rowdescription(rowdescription)
                # Print the rowdescription and the corresponding language code
                print(f"Description: {rowdescription} → Language: {language}")

def process_directory(directory_path):
    """
    Process all CSV files in the given directory and list languages found in each file.
    """
    # Get all CSV files in the directory
    files = glob.glob(os.path.join(directory_path, "*.csv"))
    
    if not files:
        print(f"No CSV files found in directory: {directory_path}")
        return

    for file_path in files:
        print(f"Processing file: {os.path.basename(file_path)}")
        process_file(file_path)
        print("-" * 40)  # Separator for readability

# Example usage
directory_path = "./input/monthly_files"  # Replace with the path to your directory
process_directory(directory_path)