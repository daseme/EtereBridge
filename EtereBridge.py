import os
import sys
import logging
import csv
import pandas as pd
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook  # Add this import
import json  # Add this import
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass, field
from tqdm import tqdm
from config_manager import config_manager
from file_processor import FileProcessor

@dataclass
class ProcessingResult:
    """Tracks the result of processing a single file."""
    filename: str
    success: bool
    error_message: Optional[str] = None
    warnings: List[str] = field(default_factory=list)
    metrics: Dict = field(default_factory=dict)
    output_file: Optional[str] = None

class ProcessingError(Exception):
    """Custom exception for processing-related errors."""
    pass

class EtereBridge:
    """Enhanced file processor with error recovery and progress tracking."""
    
    def __init__(self):
        """Initialize the EtereBridge processor."""
        self.config = config_manager.get_config()
        self.log_file = config_manager.setup_logging()
        self.results: List[ProcessingResult] = []
        
        # Initialize FileProcessor
        self.file_processor = FileProcessor(self.config)
        
    def print_header(self):
        """Display a welcome header with basic instructions."""
        header = """
╔════════════════════════════════════════════════════════════════════════════╗
║                        Excel File Processing Tool                           ║
╚════════════════════════════════════════════════════════════════════════════╝

This tool helps you process and transform Excel files according to specified formats.
Follow the prompts below to begin processing your files.

Version: 2.0
Log File: {log_file}
        """.format(log_file=self.log_file)
        print(header)

    def list_files(self) -> List[str]:
        """List all available files in the input directory."""
        files = [f for f in os.listdir(self.config.paths.input_dir) 
                if f.endswith('.csv')]
        if not files:
            print("\n❌ No CSV files found in the input directory:", 
                  self.config.paths.input_dir)
            print("Please add your CSV files to this directory and try again.")
            sys.exit(1)
        return files

    def select_processing_mode(self) -> str:
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

    def get_worldlink_defaults(self) -> Dict:
        """Return default values for WorldLink orders."""
        return {
            "billing_type": "Broadcast",
            "revenue_type": "Direct Response Sales",
            "agency_flag": "Agency",
            "sales_person": "House",
            "agency_fee": 0.15,  # Standard 15%
            "type": "COM",
            "affidavit": "Y",
            "is_worldlink": True  # Flag to identify WorldLink orders
        }

    def select_input_file(self, files: List[str]) -> Optional[str]:
        """Prompt the user to select a file from the input directory."""
        print("\n" + "-"*80)
        print("File Selection".center(80))
        print("-"*80)
        print("\nAvailable files for processing:")
        
        # Create two columns if there are many files
        mid_point = (len(files) + 1) // 2
        for i, filename in enumerate(files, 1):
            line = f"  [{i:2d}] {filename}"
            if i <= mid_point and i + mid_point <= len(files):
                second_file = files[i + mid_point - 1]
                second_item = f"  [{i + mid_point:2d}] {second_file}"
                print(f"{line:<40} {second_item}")
            else:
                print(line)
        
        while True:
            try:
                choice = input("\nEnter the number of the file you want to process (or 'q' to quit): ").strip()
                if choice.lower() == 'q':
                    print("\nExiting program...")
                    sys.exit(0)
                
                choice = int(choice)
                if 1 <= choice <= len(files):
                    selected_file = files[choice - 1]
                    print(f"\n✅ Selected: {selected_file}")
                    return os.path.join(self.config.paths.input_dir, selected_file)
                else:
                    print(f"❌ Please enter a number between 1 and {len(files)}")
            except ValueError:
                print("❌ Please enter a valid number or 'q' to quit")

    def extract_header_values(self, file_path: str) -> Tuple[str, str]:
        """Extract header values from first section of CSV."""
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                lines = file.readlines()[:2]
                
                header_row = [x.strip() for x in lines[0].split(',')]
                value_row = next(csv.reader([lines[1]]))
                
                header_dict = dict(zip(header_row, value_row))
                
                text_box_180 = header_dict.get('Textbox180', '').strip()
                text_box_171 = header_dict.get('Textbox171', '').strip()
                
                logging.info(f"Header values found - TextBox180: '{text_box_180}', TextBox171: '{text_box_171}'")
                
                return text_box_180, text_box_171
                
        except Exception as e:
            logging.error(f"Error reading header: {e}")
            return '', ''

    def generate_billcode(self, text_box_180: str, text_box_171: str) -> str:
        """Combine Textbox180 and Textbox171 for billcode."""
        if text_box_180 and text_box_171:
            return f"{text_box_180}:{text_box_171}"
        elif text_box_171:
            return text_box_171
        elif text_box_180:
            return text_box_180
        return ''

    def prompt_for_user_inputs(self) -> Dict:
        """Prompt the user for processing parameters."""
        print("\n" + "-"*80)
        print("Additional Information Needed".center(80))
        print("-"*80)
        
        # Get Sales Person
        sales_people = self.config.sales_people
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
            if billing_input in ['C', 'B']:
                billing_type = "Calendar" if billing_input == 'C' else "Broadcast"
                break
            print("❌ Please enter 'C' for Calendar or 'B' for Broadcast")
        
        # Revenue Type
        print("\n3. Revenue Type:")
        print("   [B] Branded Content")
        print("   [D] Direct Response Sales")
        print("   [I] Internal Ad Sales")
        print("   [P] Paid Programming")
        while True:
            revenue_input = input("\nSelect revenue type (B/D/I/P): ").strip().upper()
            if revenue_input in ['B', 'D', 'I', 'P']:
                revenue_types = {
                    'B': "Branded Content",
                    'D': "Direct Response Sales",
                    'I': "Internal Ad Sales",
                    'P': "Paid Programming"
                }
                revenue_type = revenue_types[revenue_input]
                break
            print("❌ Please enter 'B', 'D', 'I', or 'P'")
        
        # Agency Type and Fee
        print("\n4. Order Type:")
        print("   [A] Agency")
        print("   [N] Non-Agency")
        print("   [T] Trade")
        
        agency_fee = None
        while True:
            agency_input = input("\nSelect order type (A/N/T): ").strip().upper()
            if agency_input in ['A', 'N', 'T']:
                agency_types = {
                    'A': "Agency",
                    'N': "Non-Agency",
                    'T': "Trade"
                }
                agency_flag = agency_types[agency_input]
                
                if agency_input == 'A':
                    print("\n5. Agency Fee Type:")
                    print("   [S] Standard (15%)")
                    print("   [C] Custom")
                    while True:
                        fee_type = input("\nSelect fee type (S/C): ").strip().upper()
                        if fee_type == 'S':
                            agency_fee = 0.15
                            break
                        elif fee_type == 'C':
                            while True:
                                try:
                                    custom_fee = float(input("\nEnter custom fee percentage (without % symbol): "))
                                    if 0 <= custom_fee <= 100:
                                        agency_fee = custom_fee / 100
                                        break
                                    print("❌ Please enter a percentage between 0 and 100")
                                except ValueError:
                                    print("❌ Please enter a valid number")
                            break
                        print("❌ Please enter 'S' for Standard or 'C' for Custom")
                break
            print("❌ Please enter 'A' for Agency, 'N' for Non-Agency, or 'T' for Trade")
        
        print("\n✅ Information collected successfully!")
        return {
            "billing_type": billing_type,
            "revenue_type": revenue_type,
            "agency_flag": agency_flag,
            "sales_person": sales_person,
            "agency_fee": agency_fee
        }
    
    def prompt_for_language(self) -> str:
        """Prompt the user to select a language from the configured options."""
        print("\n" + "-"*80)
        print("Language Selection".center(80))
        print("-"*80)
        
        print("\nAvailable language options:")
        for idx, lang in enumerate(self.config.language_options, 1):
            print(f"   [{idx}] {lang}")
        
        while True:
            try:
                choice = input("\nSelect language (enter number): ").strip()
                if choice.lower() == 'q':
                    print("\nExiting program...")
                    sys.exit(0)
                
                choice = int(choice)
                if 1 <= choice <= len(self.config.language_options):
                    selected_lang = self.config.language_options[choice - 1]
                    print(f"\n✅ Selected: {selected_lang}")
                    return selected_lang
                else:
                    print(f"❌ Please enter a number between 1 and {len(self.config.language_options)}")
            except ValueError:
                print("❌ Please enter a valid number or 'q' to quit")

    def prompt_for_type(self) -> str:
        """Prompt the user to select a type from the configured options."""
        print("\n" + "-"*80)
        print("Type Selection".center(80))
        print("-"*80)
        
        print("\nAvailable type options:")
        for idx, type_opt in enumerate(self.config.type_options, 1):
            print(f"   [{idx}] {type_opt}")
        
        while True:
            try:
                choice = input("\nSelect type (enter number): ").strip()
                if choice.lower() == 'q':
                    print("\nExiting program...")
                    sys.exit(0)
                
                choice = int(choice)
                if 1 <= choice <= len(self.config.type_options):
                    selected_type = self.config.type_options[choice - 1]
                    print(f"\n✅ Selected: {selected_type}")
                    return selected_type
                else:
                    print(f"❌ Please enter a number between 1 and {len(self.config.type_options)}")
            except ValueError:
                print("❌ Please enter a valid number or 'q' to quit")

    def prompt_for_affidavit(self) -> str:
        """Prompt the user to select 'Y' or 'N' for the Affidavit column."""
        print("\n" + "-"*80)
        print("Affidavit Selection".center(80))
        print("-"*80)
        
        while True:
            affidavit_input = input("\nIs this an affidavit? (Y/N): ").strip().upper()
            if affidavit_input in ['Y', 'N']:
                print(f"\n✅ Selected: {affidavit_input}")
                return affidavit_input
            else:
                print("❌ Please enter 'Y' for Yes or 'N' for No")

    def verify_languages(self, df: pd.DataFrame, language_info: Tuple[Dict[str, int], pd.Series]) -> pd.Series:
        """
        Show detected languages and verify accuracy.
        """
        detected_counts, row_languages = language_info
        
        print("\n" + "-"*80)
        print("Language Detection Results".center(80))
        print("-"*80)
        
        # Show what was found
        for lang_code, count in detected_counts.items():
            lang_name = next((k for k, v in self.file_processor.language_mapping.items() 
                            if v == lang_code and k != 'default'), "English")
            print(f"   • {lang_name} ({lang_code}): {count} entries")
        
        # Quick verification of first few rows of each language
        print("\nSample entries:")
        for lang_code in detected_counts:
            rows = df[row_languages == lang_code]
            if not rows.empty:
                print(f"\n{lang_code}:")
                samples = rows['rowdescription'].head(2)
                for desc in samples:
                    print(f"   • {desc}")
        
        # Only ask for verification if something seems off
        if len(detected_counts) > 1:  # Multiple languages detected
            print("\nDoes this look correct? (Y/N)")
            if input().strip().lower() == 'n':
                # Show available options
                print("\nAvailable language codes:")
                for idx, code in enumerate(self.config.language_options, 1):
                    print(f"   [{idx}] {code}")
                
                # Allow fixes
                while True:
                    print("\nEnter row number to change language, or press Enter to continue")
                    row_input = input().strip()
                    if not row_input:
                        break
                    
                    try:
                        row_idx = int(row_input)
                        if 0 <= row_idx < len(df):
                            print(f"Current: {df.iloc[row_idx]['rowdescription']}")
                            print(f"Language: {row_languages.iloc[row_idx]}")
                            new_lang = input("Enter new language code: ").strip().upper()
                            if new_lang in self.config.language_options:
                                row_languages.iloc[row_idx] = new_lang
                    except ValueError:
                        print("Invalid input, try again")
        
        return row_languages

    def apply_user_inputs(self, df: pd.DataFrame, billing_type: str, revenue_type: str, 
                            agency_flag: str, sales_person: str, agency_fee: Optional[float],
                            language: Dict, type_: str, affidavit: str, is_worldlink: bool = False) -> pd.DataFrame:
        """Apply user input to the appropriate columns in the DataFrame."""
        try:
            logging.info("Applying user inputs to DataFrame...")
            
            # Add user input columns
            df['Billing Type'] = billing_type
            df['Revenue Type'] = revenue_type
            df['Agency?'] = agency_flag
            df['Sales Person'] = sales_person
            df['Lang.'] = df.index.map(language)
            df['Type'] = type_
            df['Affidavit?'] = affidavit

            # Handle WorldLink specific processing
            if is_worldlink:
                logging.info("Processing WorldLink order specific requirements...")
                # Ensure Market column exists before copying
                if 'Market' in df.columns:
                    logging.info("Copying Market data to Makegood column")
                    # Create Makegood column if it doesn't exist
                    if 'Make Good' not in df.columns:
                        df['Make Good'] = None
                    # Copy Market data to Makegood
                    df['Make Good'] = df['Market']
                    logging.info("Successfully copied Market data to Make Good")
                else:
                    logging.warning("Market column not found in WorldLink order - cannot copy to Make Good")

            # Handle agency fees
            if agency_flag == "Agency" and agency_fee is not None:
                try:
                    gross_rates = df['Gross Rate'].str.replace('$', '').str.replace(',', '').astype(float)
                    df['Broker Fees'] = gross_rates * agency_fee
                    df['Broker Fees'] = df['Broker Fees'].map('${:,.2f}'.format)
                    logging.info(f"Successfully calculated broker fees using {agency_fee:.1%} rate")
                except Exception as e:
                    logging.error(f"Error calculating broker fees: {str(e)}")
                    df['Broker Fees'] = None
            else:
                df['Broker Fees'] = None
                logging.info("No broker fees applied (non-agency or no fee specified)")

            # Ensure all required columns exist
            logging.info("Ensuring all required columns exist...")
            for col in self.config.final_columns:
                if col not in df.columns:
                    logging.info(f"Adding missing column: {col}")
                    df[col] = None
            
            # Reorder columns according to config
            logging.info("Reordering columns according to configuration...")
            try:
                df = df[self.config.final_columns]
            except KeyError as e:
                missing_cols = [col for col in self.config.final_columns if col not in df.columns]
                logging.error(f"Missing columns: {missing_cols}")
                raise KeyError(f"Missing required columns: {missing_cols}")
            
            logging.info("Successfully applied user inputs!")
            return df
                
        except Exception as e:
            logging.error(f"Error applying user inputs: {str(e)}")
            raise


    def save_to_excel(self, df: pd.DataFrame, output_path: str, agency_fee: Optional[float] = 0.15):
        try:
            # Get template path using correct attribute name
            template_path = self.config.paths.template_path
            logging.info(f"Loading template from: {template_path}")
            workbook = load_workbook(template_path, data_only=False)
            sheet = workbook.active
            
            # Get column indices from config
            columns = self.config.final_columns
            
            # Extract formulas and formatting from the second row of the template
            template_formulas = {}
            template_formatting = {}
            for col in range(1, len(columns) + 1):
                cell = sheet.cell(row=2, column=col)
                if cell.value and str(cell.value).startswith('='):  # Check if it's a formula
                    template_formulas[col] = cell.value
                # Store formatting (style, number format, etc.)
                template_formatting[col] = {
                    'style': cell.style,
                    'number_format': cell.number_format,
                    'border': cell.border.copy(),  # Create a copy of the border
                    'fill': cell.fill.copy(),      # Create a copy of the fill
                    'font': cell.font.copy(),      # Create a copy of the font
                    'alignment': cell.alignment.copy()  # Create a copy of the alignment
                }
            
            # Write headers
            for col_num, column_title in enumerate(columns, 1):
                sheet.cell(row=1, column=col_num, value=column_title)
            
            # Write data and apply formulas/formatting
            for row_num, row_data in enumerate(df.values, 2):
                for col_num, cell_value in enumerate(row_data, 1):
                    cell = sheet.cell(row=row_num, column=col_num)
                    if col_num in template_formulas:
                        # Apply the formula to the new row
                        formula = template_formulas[col_num]
                        formula = formula.replace('2', str(row_num))  # Adjust row reference
                        cell.value = formula
                    else:
                        cell.value = cell_value
                    
                    # Apply formatting from the template's second row
                    if col_num in template_formatting:
                        cell.style = template_formatting[col_num]['style']
                        cell.number_format = template_formatting[col_num]['number_format']
                        cell.border = template_formatting[col_num]['border']
                        cell.fill = template_formatting[col_num]['fill']
                        cell.font = template_formatting[col_num]['font']
                        cell.alignment = template_formatting[col_num]['alignment']
            
            # Fix the Month column (column S)
            month_col = columns.index('Month') + 1  # Get the column index for 'Month'
            for row_num in range(2, len(df) + 2):  # Iterate over all rows
                air_date_cell = sheet.cell(row=row_num, column=columns.index('Air Date') + 1)
                if air_date_cell.value:  # Check if Air Date is valid
                    try:
                        # Calculate the month based on Air Date
                        air_date = pd.to_datetime(air_date_cell.value)
                        month_value = air_date.strftime('%b-%y')  # Format as 'Dec-24'
                        sheet.cell(row=row_num, column=month_col, value=month_value)
                    except Exception as e:
                        logging.warning(f"Error calculating month for row {row_num}: {e}")
                        sheet.cell(row=row_num, column=month_col, value="Invalid Date")
                else:
                    sheet.cell(row=row_num, column=month_col, value="No Date")
            
            # Set the Priority column (column U) to 4 for all rows
            priority_col = columns.index('Priority') + 1  # Get the column index for 'Priority'
            for row_num in range(2, len(df) + 2):  # Iterate over all rows
                sheet.cell(row=row_num, column=priority_col, value=4)
            
            # Remove excess rows if the template has more rows than the CSV
            if sheet.max_row > len(df) + 1:
                sheet.delete_rows(len(df) + 2, sheet.max_row - (len(df) + 1))
            
            # Ensure output directory exists
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            workbook.save(output_path)
            logging.info("Excel file saved successfully with formulas, formatting, and Priority Number preserved")
            
        except Exception as e:
            logging.error(f"Error saving to Excel: {str(e)}")
            raise

    def generate_processing_summary(self, df: pd.DataFrame, input_file: str, output_file: str, user_inputs: Dict) -> Dict:
        """Generate comprehensive summary of the processing results."""
        try:
            # Convert date column to datetime
            df['Air Date'] = pd.to_datetime(df['Air Date'])
            
            # Convert Gross Rate to numeric for calculations
            gross_values = pd.to_numeric(
                df['Gross Rate'].str.replace('$', '').str.replace(',', ''),
                errors='coerce'
            ).fillna(0)
            
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
                    "total_gross_value": float(gross_values.sum()),
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
                }
            }
            
            logging.info(f"Generated summary for {input_file}")
            return summary
            
        except Exception as e:
            logging.error(f"Error generating summary: {e}")
            raise

    def process_file(self, file_path: str, user_inputs: Optional[Dict] = None) -> ProcessingResult:
        """Process a single input file with enhanced error handling."""
        filename = os.path.basename(file_path)
        logging.info(f"###### Starting processing of {filename} ######")
        
        try:
            # Extract header values
            logging.info("Extracting header values...")
            text_box_180, text_box_171 = self.extract_header_values(file_path)
            
            # Load and clean data using FileProcessor
            logging.info("Loading and cleaning data...")
            df = self.file_processor.load_and_clean_data(file_path)
            
            # Detect and verify languages using FileProcessor
            logging.info("Detecting languages in data...")
            detected_counts, row_languages = self.file_processor.detect_languages(df)
            logging.info(f"Detected language counts: {detected_counts}")
            
            # Apply transformations using FileProcessor
            logging.info("Applying transformations...")
            df = self.file_processor.apply_transformations(df, text_box_180, text_box_171)
            
            # Get user inputs if not provided
            if user_inputs is None:
                logging.info("Collecting user inputs...")
                user_inputs = self.prompt_for_user_inputs()
                # Add type selection
                type_ = self.prompt_for_type()
                user_inputs['type'] = type_
                # Add affidavit selection
                affidavit = self.prompt_for_affidavit()
                user_inputs['affidavit'] = affidavit
                user_inputs['is_worldlink'] = False  # Default to False for manual input
            
            # Convert row_languages (Series) to a dictionary
            language_dict = row_languages.to_dict() if not row_languages.empty else {}
            
            # Add the language dictionary to user_inputs
            user_inputs['language'] = language_dict
            
            # Apply user inputs with row-specific languages
            logging.info("Applying user inputs...")
            df = self.apply_user_inputs(
                df,
                billing_type=user_inputs['billing_type'],
                revenue_type=user_inputs['revenue_type'],
                agency_flag=user_inputs['agency_flag'],
                sales_person=user_inputs['sales_person'],
                agency_fee=user_inputs['agency_fee'],
                language=language_dict,
                type_=user_inputs['type'],
                affidavit=user_inputs['affidavit'],
                is_worldlink=user_inputs.get('is_worldlink', False)
            )
            
            # Save output file
            logging.info("Saving output file...")
            # Generate output filename
            output_filename = f"processed_{os.path.splitext(filename)[0]}.xlsx"
            output_path = os.path.join(self.config.paths.output_dir, output_filename)

            
            # Save using save_to_excel instead of save_output_file
            self.save_to_excel(df, output_path, user_inputs.get('agency_fee'))
            
            # Generate and save summary
            logging.info("Generating processing summary...")
            summary = self.generate_processing_summary(df, file_path, output_path, user_inputs)
            
            # Add language detection info to summary
            language_distribution = row_languages.value_counts().to_dict() if not row_languages.empty else {}
            summary["language_info"] = {
                "detected_languages": detected_counts,
                "language_distribution": language_distribution
            }
            
            # Add WorldLink status to summary if applicable
            if user_inputs.get('is_worldlink', False):
                summary["processing_info"]["worldlink_order"] = True
                if 'Market' in df.columns:
                    summary["processing_info"]["market_to_makegood"] = "copied"
                else:
                    summary["processing_info"]["market_to_makegood"] = "failed - Market column not found"
            
            return ProcessingResult(
                filename=filename,
                success=True,
                output_file=output_path,
                metrics=summary
            )
            
        except FileNotFoundError as e:
            error_msg = f"File not found: {filename}"
            logging.error(error_msg)
            return ProcessingResult(
                filename=filename,
                success=False,
                error_message=error_msg
            )
        except pd.errors.EmptyDataError as e:
            error_msg = f"File is empty: {filename}"
            logging.error(error_msg)
            return ProcessingResult(
                filename=filename,
                success=False,
                error_message=error_msg
            )
        except ProcessingError as e:
            error_msg = f"Processing error in {filename}: {str(e)}"
            logging.error(error_msg)
            return ProcessingResult(
                filename=filename,
                success=False,
                error_message=error_msg
            )
        except Exception as e:
            error_msg = f"Unexpected error processing {filename}: {str(e)}"
            logging.error(error_msg, exc_info=True)
            return ProcessingResult(
                filename=filename,
                success=False,
                error_message=error_msg
            )

    def process_batch(self, files: List[str], show_progress: bool = True) -> Dict[str, List[ProcessingResult]]:
        """Process multiple files with progress tracking and error recovery."""
        successful = []
        failed = []
        
        # First, check if this is a WorldLink batch
        print("\n" + "-"*80)
        print("Batch Processing Setup".center(80))
        print("-"*80)
        
        is_worldlink = input("\nIs this a batch of WorldLink orders? (Y/N): ").strip().lower() == 'y'
        
        base_user_inputs = None
        if is_worldlink:
            print("\nUsing WorldLink default settings...")
            base_user_inputs = self.get_worldlink_defaults()
            logging.info("Using WorldLink default settings for batch processing")
        else:
            # Existing logic for non-WorldLink batches
            shared_inputs = input("\nDo all files in this batch share the same user inputs? (Y/N): ").strip().lower()
            if shared_inputs == 'y':
                print("\nCollecting shared user inputs for the batch...")
                base_user_inputs = self.prompt_for_user_inputs()
                # Add type selection
                type_ = self.prompt_for_type()
                base_user_inputs['type'] = type_
                # Add affidavit selection
                affidavit = self.prompt_for_affidavit()
                base_user_inputs['affidavit'] = affidavit
        
        files_iter = tqdm(files, desc="Processing files") if show_progress else files
        
        for file_path in files_iter:
            try:
                # Load the data first to detect language using FileProcessor
                df = self.file_processor.load_and_clean_data(file_path)
                
                # Detect languages in this file using FileProcessor
                detected_languages = self.file_processor.detect_languages(df)
                
                if base_user_inputs:
                    # Create a copy of base inputs for this file
                    file_inputs = base_user_inputs.copy()
                else:
                    # Get new inputs for this file
                    file_inputs = self.prompt_for_user_inputs()
                    type_ = self.prompt_for_type()
                    file_inputs['type'] = type_
                    affidavit = self.prompt_for_affidavit()
                    file_inputs['affidavit'] = affidavit
                
                # Always detect and verify language for each file
                print(f"\nProcessing file: {os.path.basename(file_path)}")
                primary_language = self.verify_languages(df, detected_languages)
                file_inputs['language'] = primary_language
                
                # Process the file with complete inputs
                result = self.process_file(file_path, file_inputs)
                
                if result.success:
                    successful.append(result)
                else:
                    failed.append(result)
                
                # Save interim results
                self._save_interim_results(successful, failed)
                
            except Exception as e:
                logging.error(f"Error processing {file_path}: {str(e)}")
                failed.append(ProcessingResult(
                    filename=os.path.basename(file_path),
                    success=False,
                    error_message=str(e)
                ))
        
        self._display_batch_summary(successful, failed)
        return {"successful": successful, "failed": failed}

    def _save_interim_results(self, successful: List[ProcessingResult], failed: List[ProcessingResult]):
        """Save interim results to protect against crashes."""
        interim_file = Path(self.config.paths.output_dir) / 'interim_results.json'
        
        # Convert ProcessingResult objects to dictionaries
        results = {
            "timestamp": datetime.now().isoformat(),
            "successful": [],
            "failed": []
        }
        
        # Convert successful results
        for result in successful:
            result_dict = vars(result)
            # Convert Series to dict if present in metrics
            if "metrics" in result_dict and "language_distribution" in result_dict["metrics"]:
                result_dict["metrics"]["language_distribution"] = result_dict["metrics"]["language_distribution"].to_dict()
            results["successful"].append(result_dict)
        
        # Convert failed results
        for result in failed:
            results["failed"].append(vars(result))
        
        # Save to JSON
        with open(interim_file, 'w') as f:
            json.dump(results, f, indent=2)

    def _display_batch_summary(self, successful: List[ProcessingResult], 
                             failed: List[ProcessingResult]):
        """Display a user-friendly summary of batch processing results."""
        print("\n" + "="*80)
        print("Batch Processing Summary".center(80))
        print("="*80)
        
        # Success/failure counts
        total = len(successful) + len(failed)
        success_rate = (len(successful) / total * 100) if total > 0 else 0
        
        print(f"\nTotal files processed: {total}")
        print(f"Successfully processed: {len(successful)} ({success_rate:.1f}%)")
        print(f"Failed to process: {len(failed)}")
        
        # Display failures
        if failed:
            print("\nFailed Files:")
            for result in failed:
                print(f"❌ {result.filename}")
                print(f"   Error: {result.error_message}")
        
        # Display warnings
        if any(r.warnings for r in successful):
            print("\nWarnings:")
            for result in successful:
                if result.warnings:
                    print(f"⚠️ {result.filename}:")
                    for warning in result.warnings:
                        print(f"   - {warning}")

        # Display output locations
        if successful:
            print("\nProcessed Files:")
            for result in successful:
                print(f"✅ {result.filename} -> {result.output_file}")

        print(f"\nDetailed logs available at: {self.log_file}")

    def main(self):
        """Main function to control the flow of the program."""
        self.print_header()
        
        try:
            files = self.list_files()
            if not files:
                print("No files found to process. Please add files and try again.")
                return

            choice = self.select_processing_mode()

            if choice == 'A':
                print("\n🔄 Processing all files automatically...")
                file_paths = [os.path.join(self.config.paths.input_dir, f) for f in files]
                results = self.process_batch(file_paths)
                
            elif choice == 'S':
                while True:
                    file_path = self.select_input_file(files)
                    if file_path:
                        results = self.process_batch([file_path], show_progress=False)
                    
                    print("\n" + "-"*80)
                    cont = input("\nWould you like to process another file? (Y/N): ").strip().lower()
                    if cont != 'y':
                        print("\n✅ Processing complete! Thank you for using the tool.")
                        break

        except KeyboardInterrupt:
            print("\n\nProgram interrupted by user. Saving interim results...")
            self._save_interim_results(self.results, [])
            print("Interim results saved. Exiting...")
            sys.exit(0)
        except Exception as e:
            logging.error(f"Unexpected error: {str(e)}")
            print(f"\n❌ An unexpected error occurred. Please check the log file: {self.log_file}")
            sys.exit(1)

if __name__ == "__main__":
    processor = EtereBridge()
    processor.main()