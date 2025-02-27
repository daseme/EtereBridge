import os
import sys
import logging
import csv
from copy import copy
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
from user_interface import collect_user_inputs, verify_languages  # Updated import

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
        files = [f for f in os.listdir(self.config.paths.input_dir) if f.endswith('.csv')]
        if not files:
            print("\n❌ No CSV files found in the input directory:", self.config.paths.input_dir)
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

    def apply_user_inputs(self, df: pd.DataFrame, billing_type: str, revenue_type: str, 
                            agency_flag: str, sales_person: str, agency_fee: Optional[float],
                            language: Dict, affidavit: str, estimate: str, contract: str, 
                            is_worldlink: bool = False) -> pd.DataFrame:
        """
        Apply user input values to the DataFrame.
        
        Adds columns for billing type, revenue type, agency flag, sales person,
        language, affidavit, and estimate, then handles WorldLink-specific
        processing and broker fees. Also computes the Type column automatically:
        - If the Gross Rate is blank or zero, Type is 'BNS'
        - Otherwise, Type is 'COM'
        Ensures all required columns exist and orders them according to configuration.
        """
        try:
            logging.info("Applying user inputs to DataFrame...")
            
            # Add user input columns
            df['Billing Type'] = billing_type
            df['Revenue Type'] = revenue_type
            df['Agency?'] = agency_flag
            df['Sales Person'] = sales_person
            df['Lang.'] = df.index.map(language)
            df['Affidavit?'] = affidavit
            
            # Add Estimate and Contract columns from user input
            df['Estimate'] = estimate
            df['Contract'] = contract   # <-- New: write contract number
            
            # Compute Type automatically from Gross Rate on a per-row basis.
            def compute_type(row):
                try:
                    # Get the Gross Rate value (assumed formatted like "$0.00")
                    value = row.get("Gross Rate", "$0")
                    num = float(value.replace('$','').replace(',', ''))
                    if num == 0:
                        return "BNS"
                    else:
                        return "COM"
                except Exception as e:
                    logging.warning(f"Error computing type for row: {e}")
                    return "BNS"
            
            df["Type"] = df.apply(compute_type, axis=1)
            
            # Handle WorldLink-specific processing
            if is_worldlink:
                logging.info("Processing WorldLink order specific requirements...")
                if 'Market' in df.columns:
                    logging.info("Copying Market data to Makegood column")
                    if 'Make Good' not in df.columns:
                        df['Make Good'] = None
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
            
            # Reorder columns according to configuration
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

                    
        except Exception as e:
            logging.error(f"Error applying user inputs: {str(e)}")
            raise


    def save_to_excel(self, df: pd.DataFrame, output_path: str, agency_fee: Optional[float] = 0.15):
        try:
            template_path = self.config.paths.template_path
            logging.info(f"Loading template from: {template_path}")
            workbook = load_workbook(template_path, data_only=False)
            sheet = workbook.active
            
            columns = self.config.final_columns
            
            # --- 1) Extract formulas/formatting from the template's second row ---
            template_formulas = {}
            template_formatting = {}
            for col in range(1, len(columns) + 1):
                cell = sheet.cell(row=2, column=col)
                # If there's a formula in row 2, store it for later
                if cell.value and str(cell.value).startswith('='):
                    template_formulas[col] = cell.value
                # Save styling info
                template_formatting[col] = {
                    'style': cell.style,
                    'number_format': cell.number_format,
                    'border': copy(cell.border),
                    'fill': copy(cell.fill),
                    'font': copy(cell.font),
                    'alignment': copy(cell.alignment)
                }
            
            # --- 2) Write headers in row 1 ---
            for col_num, column_title in enumerate(columns, start=1):
                sheet.cell(row=1, column=col_num, value=column_title)
            
            # --- 3) Write data (row 2 onward) and apply template formatting ---
            for row_num, row_data in enumerate(df.values, start=2):
                for col_num, cell_value in enumerate(row_data, start=1):
                    cell = sheet.cell(row=row_num, column=col_num)
                    
                    # If there's a formula in the template for this column, use it
                    if col_num in template_formulas:
                        formula = template_formulas[col_num]
                        # Replace row reference '2' with the actual row number
                        formula = formula.replace('2', str(row_num))
                        cell.value = formula
                    else:
                        cell.value = cell_value
                    
                    # Apply stored formatting (style, number_format, etc.)
                    if col_num in template_formatting:
                        fmt = template_formatting[col_num]
                        cell.style = fmt['style']
                        cell.number_format = fmt['number_format']
                        cell.border = fmt['border']
                        cell.fill = fmt['fill']
                        cell.font = fmt['font']
                        cell.alignment = fmt['alignment']
            
            # --- 4) Format the Month column if needed (unchanged if it's working fine) ---
            if "Month" in columns:
                month_col = columns.index("Month") + 1
                air_date_idx = columns.index("Air Date") + 1 if "Air Date" in columns else None
                for row_num in range(2, len(df) + 2):
                    if air_date_idx:
                        air_date_val = sheet.cell(row=row_num, column=air_date_idx).value
                        if air_date_val:
                            try:
                                dt = pd.to_datetime(air_date_val)
                                month_cell = sheet.cell(row=row_num, column=month_col, value=dt)
                                # Example: keep "m/d/yyyy" if that was already working
                                month_cell.number_format = "m/d/yyyy"
                            except Exception as e:
                                logging.warning(f"Error formatting Month row {row_num}: {e}")
                                sheet.cell(row=row_num, column=month_col, value="Invalid Date")
                        else:
                            sheet.cell(row=row_num, column=month_col, value="No Date")
            
            # --- 5) Format Air Date with 2-digit year (m/d/yy) ---
            if "Air Date" in columns:
                air_date_col = columns.index("Air Date") + 1
                for row_num in range(2, len(df) + 2):
                    cell = sheet.cell(row=row_num, column=air_date_col)
                    if cell.value:
                        try:
                            dt = pd.to_datetime(cell.value)
                            cell.value = dt
                            cell.number_format = "m/d/yy"  # 2-digit year
                        except Exception as e:
                            logging.warning(f"Error formatting Air Date row {row_num}: {e}")
            
            # --- 6) Format End Date with 2-digit year (m/d/yy) ---
            if "End Date" in columns:
                end_date_col = columns.index("End Date") + 1
                for row_num in range(2, len(df) + 2):
                    cell = sheet.cell(row=row_num, column=end_date_col)
                    if cell.value:
                        try:
                            dt = pd.to_datetime(cell.value)
                            cell.value = dt
                            cell.number_format = "m/d/yy"  # 2-digit year
                        except Exception as e:
                            logging.warning(f"Error formatting End Date row {row_num}: {e}")
            
            # --- 7) (Optional) Set the Priority column to 4 if it exists ---
            if "Priority" in columns:
                priority_col = columns.index("Priority") + 1
                for row_num in range(2, len(df) + 2):
                    sheet.cell(row=row_num, column=priority_col, value=4)
            
            # --- 8) Remove extra rows if the template has more rows than the CSV data ---
            if sheet.max_row > len(df) + 1:
                sheet.delete_rows(len(df) + 2, sheet.max_row - (len(df) + 1))
            
            # Ensure the output directory exists
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Save the final workbook
            workbook.save(output_path)
            logging.info("Excel file saved successfully with formulas, formatting, and Priority Number preserved")
        
        except Exception as e:
            logging.error(f"Error saving to Excel: {str(e)}")
            raise

    def generate_processing_summary(self, df: pd.DataFrame, input_file: str, output_file: str, user_inputs: Dict) -> Dict:
        try:
            df['Air Date'] = pd.to_datetime(df['Air Date'])
            gross_values = pd.to_numeric(
                df['Gross Rate'].str.replace('$', '').str.replace(',', ''),
                errors='coerce'
            ).fillna(0)
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
        filename = os.path.basename(file_path)
        logging.info(f"###### Starting processing of {filename} ######")
        
        try:
            logging.info("Extracting header values...")
            text_box_180, text_box_171 = self.extract_header_values(file_path)
            logging.info("Loading and cleaning data...")
            df = self.file_processor.load_and_clean_data(file_path)
            logging.info("Detecting languages in data...")
            detected_counts, row_languages = self.file_processor.detect_languages(df)
            logging.info(f"Detected language counts: {detected_counts}")
            logging.info("Applying transformations...")
            df = self.file_processor.apply_transformations(df, text_box_180, text_box_171)
            
            if user_inputs is None:
                logging.info("Collecting user inputs...")
                user_inputs = collect_user_inputs(self.config)
                user_inputs['is_worldlink'] = False
            
            language_dict = row_languages.to_dict() if not row_languages.empty else {}
            user_inputs['language'] = language_dict
            
            logging.info("Verifying languages...")
            primary_language = verify_languages(df, (detected_counts, row_languages))
            if isinstance(primary_language, pd.Series):
                primary_language = primary_language.to_dict()
            user_inputs['language'] = primary_language

            
            user_inputs['language'] = primary_language
            
            logging.info("Applying user inputs...")
            df = self.apply_user_inputs(
                df,
                billing_type=user_inputs['billing_type'],
                revenue_type=user_inputs['revenue_type'],
                agency_flag=user_inputs['agency_flag'],
                sales_person=user_inputs['sales_person'],
                agency_fee=user_inputs['agency_fee'],
                language=language_dict,
                affidavit=user_inputs['affidavit'],
                estimate=user_inputs['estimate'],
                contract=user_inputs['contract'],
                is_worldlink=user_inputs.get('is_worldlink', False)
            )
            
            logging.info("Saving output file...")
            output_filename = f"processed_{os.path.splitext(filename)[0]}.xlsx"
            output_path = os.path.join(self.config.paths.output_dir, output_filename)
            self.save_to_excel(df, output_path, user_inputs.get('agency_fee'))
            
            logging.info("Generating processing summary...")
            summary = self.generate_processing_summary(df, file_path, output_path, user_inputs)
            language_distribution = row_languages.value_counts().to_dict() if not row_languages.empty else {}
            summary["language_info"] = {
                "detected_languages": detected_counts,
                "language_distribution": language_distribution
            }
            
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
            return ProcessingResult(filename=filename, success=False, error_message=error_msg)
        except pd.errors.EmptyDataError as e:
            error_msg = f"File is empty: {filename}"
            logging.error(error_msg)
            return ProcessingResult(filename=filename, success=False, error_message=error_msg)
        except ProcessingError as e:
            error_msg = f"Processing error in {filename}: {str(e)}"
            logging.error(error_msg)
            return ProcessingResult(filename=filename, success=False, error_message=error_msg)
        except Exception as e:
            error_msg = f"Unexpected error processing {filename}: {str(e)}"
            logging.error(error_msg, exc_info=True)
            return ProcessingResult(filename=filename, success=False, error_message=error_msg)

    def process_batch(self, files: List[str], show_progress: bool = True) -> Dict[str, List[ProcessingResult]]:
        successful = []
        failed = []
        
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
            shared_inputs = input("\nDo all files in this batch share the same user inputs? (Y/N): ").strip().lower()
            if shared_inputs == 'y':
                print("\nCollecting shared user inputs for the batch...")
                base_user_inputs = collect_user_inputs(self.config)
        
        files_iter = tqdm(files, desc="Processing files") if show_progress else files
        
        for file_path in files_iter:
            try:
                df = self.file_processor.load_and_clean_data(file_path)
                detected_languages = self.file_processor.detect_languages(df)
                
                if base_user_inputs:
                    file_inputs = base_user_inputs.copy()
                else:
                    file_inputs = collect_user_inputs(self.config)
                
                print(f"\nProcessing file: {os.path.basename(file_path)}")
                primary_language = verify_languages(df, detected_languages)
                if isinstance(primary_language, pd.Series):
                    primary_language = primary_language.to_dict()
                file_inputs['language'] = primary_language

                
                result = self.process_file(file_path, file_inputs)
                
                if result.success:
                    successful.append(result)
                else:
                    failed.append(result)
                
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
        interim_file = Path(self.config.paths.output_dir) / 'interim_results.json'
        results = {
            "timestamp": datetime.now().isoformat(),
            "successful": [],
            "failed": []
        }
        for result in successful:
            result_dict = vars(result)
            if "metrics" in result_dict and "language_distribution" in result_dict["metrics"]:
                result_dict["metrics"]["language_distribution"] = result_dict["metrics"]["language_distribution"].to_dict()
            results["successful"].append(result_dict)
        for result in failed:
            results["failed"].append(vars(result))
        with open(interim_file, 'w') as f:
            json.dump(results, f, indent=2)

    def _display_batch_summary(self, successful: List[ProcessingResult], failed: List[ProcessingResult]):
        print("\n" + "="*80)
        print("Batch Processing Summary".center(80))
        print("="*80)
        
        total = len(successful) + len(failed)
        success_rate = (len(successful) / total * 100) if total > 0 else 0
        
        print(f"\nTotal files processed: {total}")
        print(f"Successfully processed: {len(successful)} ({success_rate:.1f}%)")
        print(f"Failed to process: {len(failed)}")
        
        if failed:
            print("\nFailed Files:")
            for result in failed:
                print(f"❌ {result.filename}")
                print(f"   Error: {result.error_message}")
        
        if any(r.warnings for r in successful):
            print("\nWarnings:")
            for result in successful:
                if result.warnings:
                    print(f"⚠️ {result.filename}:")
                    for warning in result.warnings:
                        print(f"   - {warning}")
        if successful:
            print("\nProcessed Files:")
            for result in successful:
                print(f"✅ {result.filename} -> {result.output_file}")

        print(f"\nDetailed logs available at: {self.log_file}")

    def main(self):
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
