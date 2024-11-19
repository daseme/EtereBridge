import os
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import math
import sys
import csv
import json
from config_manager import config_manager



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
    files = [f for f in os.listdir(config_manager.get_config().paths.input_dir) if f.endswith('.csv')]
    if not files:
        print("\n❌ No CSV files found in the input directory:", config_manager.get_config().paths.input_dir)
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
                return os.path.join(config_manager.get_config().paths.input_dir, selected_file)
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
        
        # Update market replacements to use config
        df['Market'] = df['Market'].replace(config_manager.get_config().market_replacements)
        
        # Rest of transformations
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
    
    # Get Sales Person
    sales_people = config_manager.get_config().sales_people
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
    for col in config_manager.get_config().final_columns:
        if col not in df.columns:
            df[col] = None
    
    print("🔄 Reordering columns...")
    df = df[config_manager.get_config().final_columns]
    
    print("✅ User inputs applied successfully!")
    return df

def save_to_excel(df, template_path, output_path):
    """Save DataFrame to Excel, preserving template but removing excess rows."""
    try:
        workbook = load_workbook(template_path)
        sheet = workbook.active
        
        # Write headers and data
        for col_num, column_title in enumerate(config_manager.get_config().final_columns, 1):
            sheet.cell(row=1, column=col_num, value=column_title)

        for row_num, row_data in enumerate(df.values, 2):
            for col_num, cell_value in enumerate(row_data, 1):
                column_name = config_manager.get_config().final_columns[col_num - 1]
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
        print(f"Total Spots Processed: {summary['overall_metrics']['total_spots']:,}")
        print(f"Total Gross Value: ${summary['overall_metrics']['total_gross_value']:,.2f}")
        print(f"Average Spot Value: ${summary['overall_metrics']['average_spot_value']:,.2f}")
        print(f"Unique Programs: {summary['overall_metrics']['unique_programs']}")
        
        print(f"\nDate Range: {summary['date_range']['earliest']} to {summary['date_range']['latest']}")
        print(f"Total Days: {summary['date_range']['total_days']}")
        
        print(f"\nLength Statistics:")
        print(f"Average Length: {summary['length_statistics']['average_length']}")
        print(f"Min Length: {summary['length_statistics']['min_length']}")
        print(f"Max Length: {summary['length_statistics']['max_length']}")
        
        print(f"\nMarket Breakdown:")
        for market, count in summary['breakdowns']['markets'].items():
            print(f"  {market}: {count:,} spots")
        
        print(f"\nMedia Type Breakdown:")
        for media, count in summary['breakdowns']['media_types'].items():
            print(f"  {media}: {count:,} spots")
            
        print(f"\nSpots by Day:")
        for day, count in summary['breakdowns']['spots_by_day'].items():
            print(f"  {day}: {count:,} spots")
            
    except Exception as e:
        print(f"Error displaying summary: {str(e)}")
        raise

def generate_enhanced_processing_summary(df, input_file, output_file, user_inputs):
    """
    Generate an enhanced summary of the file processing including metadata.
    
    Args:
        df (pandas.DataFrame): The processed dataframe
        input_file (str): Path to input file
        output_file (str): Path to output file
        user_inputs (dict): Dictionary containing user inputs used for processing
    
    Returns:
        dict: Enhanced summary dictionary
    """
    # Convert date column to datetime
    df['Air Date'] = pd.to_datetime(df['Air Date'])
    
    # Convert Gross Rate to numeric for calculations
    gross_values = df['Gross Rate'].str.replace('$', '').str.replace(',', '').astype(float)
    
    # Calculate spots by day of week
    df['Day_of_Week'] = df['Air Date'].dt.day_name()
    spots_by_day = df['Day_of_Week'].value_counts().to_dict()
    
    summary = {
        "processing_info": {
            "timestamp": datetime.now().isoformat(),
            "input_file": input_file,
            "output_file": output_file,
            "user_inputs": user_inputs
        },
        "overall_metrics": {
            "total_spots": len(df),
            "total_gross_value": float(gross_values.sum()),  # Convert to float for JSON serialization
            "average_spot_value": float(gross_values.mean()),
            "unique_programs": len(df['Program'].unique()),
        },
        "date_range": {
            "earliest": df['Air Date'].min().isoformat(),
            "latest": df['Air Date'].max().isoformat(),
            "total_days": (df['Air Date'].max() - df['Air Date'].min()).days + 1
        },
        "breakdowns": {
            "markets": df['Market'].value_counts().to_dict(),
            "media_types": df['Media'].value_counts().to_dict(),
            "spots_by_day": spots_by_day,
            "programs": df['Program'].value_counts().to_dict()
        },
        "length_statistics": {
            "average_length": str(pd.to_timedelta(df['Length']).mean()),
            "min_length": str(pd.to_timedelta(df['Length']).min()),
            "max_length": str(pd.to_timedelta(df['Length']).max())
        }
    }
    
    return summary

def save_processing_summary(summary, filename_base):
    """
    Save processing summary in multiple formats.
    
    Args:
        summary (dict): The processing summary dictionary
        filename_base (str): Base filename to use for saving
    
    Returns:
        tuple: Paths to saved summary files
    """
    # Create summaries directory if it doesn't exist
    summary_dir = os.path.join(config_manager.get_config().paths.output_dir, 'summaries')
    os.makedirs(summary_dir, exist_ok=True)
    
    # Create timestamp-based subdirectory
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    summary_subdir = os.path.join(summary_dir, f"{filename_base}_{timestamp}")
    os.makedirs(summary_subdir, exist_ok=True)
    
    # Save JSON version
    json_path = os.path.join(summary_subdir, "summary.json")
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)
    
    # Create flattened version for CSV
    flat_summary = {
        "Timestamp": summary["processing_info"]["timestamp"],
        "Input File": summary["processing_info"]["input_file"],
        "Output File": summary["processing_info"]["output_file"],
        "Total Spots": summary["overall_metrics"]["total_spots"],
        "Total Gross Value": summary["overall_metrics"]["total_gross_value"],
        "Average Spot Value": summary["overall_metrics"]["average_spot_value"],
        "Unique Programs": summary["overall_metrics"]["unique_programs"],
        "Start Date": summary["date_range"]["earliest"],
        "End Date": summary["date_range"]["latest"],
        "Total Days": summary["date_range"]["total_days"],
        "Average Length": summary["length_statistics"]["average_length"],
        "Billing Type": summary["processing_info"]["user_inputs"]["billing_type"],
        "Revenue Type": summary["processing_info"]["user_inputs"]["revenue_type"],
        "Agency Type": summary["processing_info"]["user_inputs"]["agency_flag"],
        "Sales Person": summary["processing_info"]["user_inputs"]["sales_person"]
    }
    
    # Save CSV version
    csv_path = os.path.join(summary_subdir, "summary.csv")
    pd.DataFrame([flat_summary]).to_csv(csv_path, index=False)
    
    print(f"\n✅ Summary files saved:")
    print(f"   - JSON: {json_path}")
    print(f"   - CSV: {csv_path}")
    
    return json_path, csv_path

def process_file(file_path):
    """Process a single input file and save processing statistics."""
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
    
    # Create user inputs dictionary for summary
    user_inputs = {
        "billing_type": billing_type,
        "revenue_type": revenue_type,
        "agency_flag": agency_flag,
        "sales_person": sales_person
    }
    
    print("\n🔄 Applying user inputs and reordering columns...")
    df = apply_user_inputs(df, billing_type, revenue_type, agency_flag, sales_person)
    print("✅ User inputs applied!")

    # Define output file name
    filename_base = os.path.splitext(os.path.basename(file_path))[0]
    timestamp = datetime.now().strftime("%Y-%m-%d")
    output_file = os.path.join(config_manager.get_config().paths.output_dir, 
                              f"{filename_base}_Processed_{timestamp}.xlsx")
    
    print("\n🔄 Saving to Excel...")
    save_to_excel(df, config_manager.get_config().paths.template_path, output_file)
    print(f"✅ File saved successfully to: {output_file}")

    # Generate and save enhanced summary
    summary = generate_enhanced_processing_summary(
        df, 
        file_path, 
        output_file, 
        user_inputs
    )
    
    # Save summary files
    json_path, csv_path = save_processing_summary(summary, filename_base)
    
    # Display summary
    display_processing_summary(summary)


def main():
    """Main function to control the flow of the program."""
    print_header()
    
    try:
        files = list_files()
        choice = select_processing_mode()

        if choice == 'A':
            print("\n🔄 Processing all files automatically...")
            for file in files:
                file_path = os.path.join(config_manager.get_config().paths.input_dir, file)  # <- Fixed here
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