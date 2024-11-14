import os
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import math
import configparser
import sys
import csv

"""

Add Billcode In for to Transformations
- TextBox 180:TextBox171
If TextBox 180 blank then use only TextBox171
Look at the 15s/30s conversion round fx
Agency function to go to agency non-agency trade

Keyword arguments:
argument -- description
Return: return_description
"""



def load_config():
    """Load configuration from config.ini file."""
    config = configparser.ConfigParser()
    # Preserve case sensitivity of keys
    config.optionxform = str
    
    # Get the directory of the current script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, 'config.ini')
    
    # Check if config file exists
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Config file not found at: {config_path}")
    
    # Read the config file
    files_read = config.read(config_path)
    if not files_read:
        raise ValueError(f"Could not read config file at: {config_path}")
    
    try:
        # Load paths
        TEMPLATE_PATH = config['Paths']['template_path']
        INPUT_DIR = config['Paths']['input_dir']
        OUTPUT_DIR = config['Paths']['output_dir']
        
        # Convert relative paths to absolute paths based on script location
        TEMPLATE_PATH = os.path.join(script_dir, TEMPLATE_PATH)
        INPUT_DIR = os.path.join(script_dir, INPUT_DIR)
        OUTPUT_DIR = os.path.join(script_dir, OUTPUT_DIR)
        
        # Create directories if they don't exist
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.makedirs(os.path.dirname(TEMPLATE_PATH), exist_ok=True)
        os.makedirs(INPUT_DIR, exist_ok=True)
        
        # Load market replacements
        MARKET_REPLACEMENTS = dict(config['Markets'])
        
        # Load and parse final columns, handling the '#' character
        FINAL_COLUMNS = config['Columns']['final_columns'].split(',')
        # Replace the placeholder with actual '#' if needed
        FINAL_COLUMNS = [col.strip() if col.strip() != 'Number' else '#' for col in FINAL_COLUMNS]
        
        return TEMPLATE_PATH, INPUT_DIR, OUTPUT_DIR, MARKET_REPLACEMENTS, FINAL_COLUMNS
        
    except KeyError as e:
        print(f"Error: Missing section or key in config file: {e}")
        print("Please ensure your config.ini file contains all required sections and keys:")
        print("Required sections: [Paths], [Markets], [Columns]")
        print("Required keys in [Paths]: template_path, input_dir, output_dir")
        sys.exit(1)
    except Exception as e:
        print(f"Error loading configuration: {e}")
        sys.exit(1)

try:
    # Load configuration at module level
    TEMPLATE_PATH, INPUT_DIR, OUTPUT_DIR, MARKET_REPLACEMENTS, FINAL_COLUMNS = load_config()
except Exception as e:
    print(f"Failed to load configuration: {e}")
    sys.exit(1)

def print_header():
    """Display a welcome header with basic instructions."""
    print("\n" + "="*80)
    print("Excel File Processing Tool".center(80))
    print("="*80)
    print("\nThis tool helps you process and transform Excel files according to specified formats.")
    print("Follow the prompts below to begin processing your files.\n")

def round_to_nearest_30_seconds(seconds):
    """Round the given number of seconds to the nearest 30-second increment.
    
    31 seconds will round down to 30 seconds
    45 seconds will round up to 60 seconds
    15 seconds will round to 0 seconds
    75 seconds will round to 90 seconds
    """
    return round(float(seconds) / 15) * 15

def select_processing_mode():
    """Ask the user whether to process all files or select one at a time."""
    print("\n" + "-"*80)
    print("Processing Mode Selection".center(80))
    print("-"*80)
    print("\nChoose how you want to process your files:")
    print("  [A] Process all files automatically")
    print("  [S] Select and process files one at a time")
    
    while True:
        choice = input("\nYour choice (A/S): ").strip().upper()
        if choice in ['A', 'S']:
            return choice
        print("❌ Invalid choice. Please enter 'A' for all files or 'S' to select files.")

def list_files():
    """List all available files in the input directory."""
    files = [f for f in os.listdir(INPUT_DIR) if f.endswith('.csv')]
    if not files:
        print("\n❌ No CSV files found in the input directory:", INPUT_DIR)
        print("Please add your CSV files to this directory and try again.")
        sys.exit(1)
    return files

def select_input_file(files):
    """Prompt the user to select a file from the input directory."""
    print("\n" + "-"*80)
    print("File Selection".center(80))
    print("-"*80)
    print("\nAvailable files for processing:")
    
    # Calculate the maximum width needed for file names
    max_width = max(len(str(i)) + len(filename) for i, filename in enumerate(files, 1))
    
    # Create two columns if there are many files
    mid_point = (len(files) + 1) // 2
    for i, filename in enumerate(files, 1):
        # Format each line with consistent spacing
        line = f"  [{i:2d}] {filename}"
        if i <= mid_point and i + mid_point <= len(files):
            # Print two columns if there are enough files
            second_file = files[i + mid_point - 1]
            second_item = f"  [{i + mid_point:2d}] {second_file}"
            print(f"{line:<40} {second_item}")
        else:
            print(line)
    
    while True:
        try:
            choice = input("\nEnter the number of the file you want to process: ").strip()
            if choice.lower() == 'q':
                print("\nExiting program...")
                sys.exit(0)
            
            choice = int(choice)
            if 1 <= choice <= len(files):
                selected_file = files[choice - 1]
                print(f"\n✅ Selected: {selected_file}")
                return os.path.join(INPUT_DIR, selected_file)
            else:
                print(f"❌ Please enter a number between 1 and {len(files)}")
        except ValueError:
            print("❌ Please enter a valid number or 'q' to quit")

def clean_numeric(value):
    """Clean numeric strings by removing commas and decimal points."""
    if isinstance(value, str):
        return value.replace(',', '').split('.')[0]
    return value

def load_and_clean_data(file_path):
    """Load data from the selected input file and perform initial transformations."""
    try:
        df = pd.read_csv(file_path, skiprows=3)
        df = df.dropna(how='all')
        df = df[~df['IMPORTO2'].astype(str).str.contains('Textbox', na=False)]
        
        # Clean numeric fields before renaming
        df['id_contrattirighe'] = df['id_contrattirighe'].apply(clean_numeric)
        df['Textbox14'] = df['Textbox14'].apply(clean_numeric)
        
        df = df.rename(columns={
            'id_contrattirighe': 'Line',
            'Textbox14': '#',
            'duration3': 'Length',
            'IMPORTO2': 'Gross Rate',
            'nome2': 'Market',
            'dateschedule': 'Air Date',
            'airtimep': 'Program',
            'bookingcode2': 'Media'
        })
        
        df[['Time In', 'Time Out']] = df['timerange2'].str.split('-', expand=True)
        return df
        
    except Exception as e:
        print(f"❌ Error loading or cleaning data: {e}")
        return None
        
    except Exception as e:
        print(f"❌ Error loading or cleaning data: {e}")
        return None

def extract_header_values(file_path):
    """Extract header values from first section of CSV."""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            # Read first two lines
            lines = file.readlines()[:2]
            
            # For the second line, we'll use a more robust CSV parsing
            import csv
            header_row = [x.strip() for x in lines[0].split(',')]
            value_row = next(csv.reader([lines[1]]))
            
            # Map headers to values
            header_dict = dict(zip(header_row, value_row))
            
            # Get values from correct columns
            text_box_180 = header_dict.get('Textbox180', '').strip()
            # Get correct agency name from Textbox171
            text_box_171 = header_dict.get('Textbox171', '').strip()
            
            # Debug print
            print("\nHeader values found:")
            for key, value in header_dict.items():
                print(f"{key}: '{value}'")
                
            return text_box_180, text_box_171
            
    except Exception as e:
        print(f"Error reading header: {e}")
        return '', ''


def generate_billcode(text_box_180, text_box_171):
    """Combine Textbox180 and Textbox171 for billcode."""
    if text_box_180 and text_box_171:
        return f"{text_box_180}:{text_box_171}"
    elif text_box_171:
        return text_box_171
    elif text_box_180:
        return text_box_180
    return ''


def apply_transformations(df, text_box_180, text_box_171):
    """Apply transformations including billcode."""
    try:
        # Set billcode first
        billcode = generate_billcode(text_box_180, text_box_171)
        df['Bill Code'] = billcode
        
        # Rest of transformations remain the same
        df['Market'] = df['Market'].replace(MARKET_REPLACEMENTS)
        df['Gross Rate'] = df['Gross Rate'].astype(str).str.replace('$', '').str.replace(',', '')
        df['Gross Rate'] = pd.to_numeric(df['Gross Rate'], errors='coerce').fillna(0).map("${:,.2f}".format)
        df['Length'] = df['Length'].apply(round_to_nearest_30_seconds)
        df['Length'] = pd.to_timedelta(df['Length'], unit='s').apply(lambda x: str(x).split()[-1].zfill(8))
        df['Line'] = pd.to_numeric(df['Line'], errors='coerce').fillna(0).astype(int)
        df['#'] = pd.to_numeric(df['#'], errors='coerce').fillna(0).astype(int)
        
        return df
        
    except Exception as e:
        print(f"Error in transformations: {e}")
        raise

def prompt_for_user_inputs():
    """Prompt the user for values to fill in certain columns."""
    print("\n" + "-"*80)
    print("Additional Information Needed".center(80))
    print("-"*80)
    
    # Load sales people from config
    config = configparser.ConfigParser()
    config.optionxform = str
    config.read('config.ini')
    sales_people = config['Sales']['sales_people'].split(',')
    
    # Get Sales Person
    print("\n1. Sales Person:")
    for idx, person in enumerate(sales_people, 1):
        print(f"   [{idx}] {person}")
    while True:
        try:
            choice = int(input("\nSelect sales person (enter number): "))
            if 1 <= choice <= len(sales_people):
                sales_person = sales_people[choice-1]
                break
            print(f"❌ Please enter a number between 1 and {len(sales_people)}")
        except ValueError:
            print("❌ Please enter a valid number")
    
    # Billing Type
    print("\n2. Billing Type:")
    print("   [C] Calendar")
    print("   [B] Broadcast")
    while True:
        billing_input = input("\nSelect billing type (C/B): ").strip().upper()
        if billing_input == 'C':
            billing_type = "Calendar"
            break
        elif billing_input == 'B':
            billing_type = "Broadcast"
            break
        print("❌ Please enter 'C' for Calendar or 'B' for Broadcast")
    
    # Revenue Type
    print("\n3. Revenue Type:")
    print("   [B] Branded Content")
    print("   [D] Direct Response")
    print("   [I] Internal Ad Sales")
    print("   [P] Paid Programming")
    while True:
        revenue_input = input("\nSelect revenue type (B/D/I/P): ").strip().upper()
        if revenue_input == 'B':
            revenue_type = "Branded Content"
            break
        elif revenue_input == 'D':
            revenue_type = "Direct Response"
            break
        elif revenue_input == 'I':
            revenue_type = "Internal Ad Sales"
            break
        elif revenue_input == 'P':
            revenue_type = "Paid Programming"
            break
        print("❌ Please enter 'B' for Branded Content, 'D' for Direct Response, 'I' for Internal Ad Sales, or 'P' for Paid Programming")
    
    # Updated Agency Flag
    print("\n4. Order Type:")
    print("   [A] Agency")
    print("   [N] Non-Agency")
    print("   [T] Trade")
    while True:
        agency_input = input("\nSelect order type (A/N/T): ").strip().upper()
        if agency_input == 'A':
            agency_flag = "Agency"
            break
        elif agency_input == 'N':
            agency_flag = "Non-Agency"
            break
        elif agency_input == 'T':
            agency_flag = "Trade"
            break
        print("❌ Please enter 'A' for Agency, 'N' for Non-Agency, or 'T' for Trade")
    
    print("\n✅ Information collected successfully!")
    return billing_type, revenue_type, agency_flag, sales_person

def apply_user_inputs(df, billing_type, revenue_type, agency_flag, sales_person):
    """Apply user input to the appropriate columns in the DataFrame."""
    print("\n🔄 Applying user inputs to data...")
    df['Billing Type'] = billing_type
    df['Revenue Type'] = revenue_type
    df['Agency?'] = agency_flag
    df['Sales Person'] = sales_person

    print("🔄 Ensuring all required columns exist...")
    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = None
    
    print("🔄 Reordering columns...")
    df = df[FINAL_COLUMNS]
    
    print("✅ User inputs applied successfully!")
    return df

def save_to_excel(df, template_path, output_path):
    """Save DataFrame to Excel, preserving template but removing excess rows."""
    try:
        workbook = load_workbook(template_path)
        sheet = workbook.active
        
        # Write headers and data
        for col_num, column_title in enumerate(FINAL_COLUMNS, 1):
            sheet.cell(row=1, column=col_num, value=column_title)

        for row_num, row_data in enumerate(df.values, 2):
            for col_num, cell_value in enumerate(row_data, 1):
                column_name = FINAL_COLUMNS[col_num - 1]
                cell = sheet.cell(row=row_num, column=col_num, value=cell_value)
                
                if column_name == '#':
                    cell.number_format = '0'
                elif column_name == 'Length':
                    cell.number_format = 'hh:mm:ss'
                elif column_name == 'Gross Rate':
                    cell.number_format = '$#,##0.00'
        
        # Remove excess rows while preserving formulas in other columns
        last_data_row = len(df) + 1  # +1 for header row
        if sheet.max_row > last_data_row:
            sheet.delete_rows(last_data_row + 1, sheet.max_row - last_data_row)
        
        workbook.save(output_path)
        
    except Exception as e:
        print(f"\n❌ Error saving to Excel: {str(e)}")

def generate_processing_summary(df):
    """Generate summary statistics for the processed file."""
    try:
        # Convert date column to datetime
        df['Air Date'] = pd.to_datetime(df['Air Date'])
        
        # Convert Gross Rate to numeric for calculations
        gross_values = df['Gross Rate'].str.replace('$', '').str.replace(',', '').astype(float)
        
        summary = {
            "total_spots": len(df),
            "total_gross_value": gross_values.sum(),
            "markets_breakdown": df['Market'].value_counts().to_dict(),
            "media_breakdown": df['Media'].value_counts().to_dict(),
            "avg_spot_length": pd.to_timedelta(df['Length']).mean(),
            "date_range": {
                "earliest": df['Air Date'].min().strftime('%Y-%m-%d'),
                "latest": df['Air Date'].max().strftime('%Y-%m-%d')
            },
            "programs": len(df['Program'].unique()),
        }
        return summary
        
    except Exception as e:
        print(f"Error generating summary: {str(e)}")
        raise

def display_processing_summary(summary):
    """Display the processing summary in a user-friendly format."""
    try:
        print("\nProcessing Summary")
        print("-" * 80)
        
        print(f"\nOverall Statistics:")
        print(f"Total Spots Processed: {summary['total_spots']:,}")
        print(f"Total Gross Value: ${summary['total_gross_value']:,.2f}")
        print(f"Average Spot Length: {summary['avg_spot_length']}")
        print(f"Date Range: {summary['date_range']['earliest']} to {summary['date_range']['latest']}")
        print(f"Unique Programs: {summary['programs']}")
        
        print(f"\nMarket Breakdown:")
        for market, count in summary['markets_breakdown'].items():
            print(f"  {market}: {count:,} spots")
        
        print(f"\nMedia Type Breakdown:")
        for media, count in summary['media_breakdown'].items():
            print(f"  {media}: {count:,} spots")
            
    except Exception as e:
        print(f"Error displaying summary: {str(e)}")
        raise

def process_file(file_path):
    """Process a single input file and display processing statistics."""
    print("\n" + "-"*80)
    print(f"Processing: {os.path.basename(file_path)}".center(80))
    print("-"*80)
    
    # Extract values from header
    print("\n🔄 Extracting header values for TextBox180 and TextBox171...")
    text_box_180, text_box_171 = extract_header_values(file_path)
    print(f"✅ Extracted TextBox180: {text_box_180}, TextBox171: {text_box_171}")

    print("\n🔄 Loading and cleaning data...")
    df = load_and_clean_data(file_path)
    if df is None:
        print(f"❌ Failed to process {file_path}. Skipping.")
        return

    print("✅ Data loaded successfully!")
    print("\n🔄 Applying transformations...")
    df = apply_transformations(df, text_box_180, text_box_171)
    print("✅ Transformations complete!")

    # Prompt for additional user inputs
    billing_type, revenue_type, agency_flag, sales_person = prompt_for_user_inputs()
    
    print("\n🔄 Applying user inputs and reordering columns...")
    df = apply_user_inputs(df, billing_type, revenue_type, agency_flag, sales_person)
    print("✅ User inputs applied!")

    # Generate and display summary
    summary = generate_processing_summary(df)
    display_processing_summary(summary)

    # Define output file name and save the result
    filename_base = os.path.splitext(os.path.basename(file_path))[0]
    timestamp = datetime.now().strftime("%Y-%m-%d")
    output_file = os.path.join(OUTPUT_DIR, f"{filename_base}_Processed_{timestamp}.xlsx")
    
    print("\n🔄 Saving to Excel...")
    save_to_excel(df, TEMPLATE_PATH, output_file)
    print(f"✅ File saved successfully to: {output_file}")


def main():
    """Main function to control the flow of the program."""
    print_header()
    
    try:
        files = list_files()
        choice = select_processing_mode()

        if choice == 'A':
            print("\n🔄 Processing all files automatically...")
            for file in files:
                file_path = os.path.join(INPUT_DIR, file)
                process_file(file_path)
                print("\n" + "="*80)
            print("\n✅ All files processed successfully!")
            
        elif choice == 'S':
            while True:
                file_path = select_input_file(files)
                if file_path:
                    process_file(file_path)
                
                print("\n" + "-"*80)
                cont = input("\nWould you like to process another file? (Y/N): ").strip().lower()
                if cont != 'y':
                    print("\n✅ Processing complete! Thank you for using the tool.")
                    break
    
    except KeyboardInterrupt:
        print("\n\nProgram interrupted by user. Exiting...")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ An unexpected error occurred: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()