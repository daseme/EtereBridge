# file_processor.py
import pandas as pd
from typing import Dict, Tuple, Optional
import logging

class FileProcessor:
    def __init__(self, config):
        """
        Initialize the FileProcessor with configuration settings.
        
        Args:
            config: Configuration object containing paths and settings.
        """
        self.config = config
        self.language_mapping = {
            "Chinese": "M",
            "Filipino": "T",
            "Hmong": "Hm",
            "South Asian": "SA",
            "Vietnamese": "V",
            "Mandarin": "M",
            "Cantonese": "C",
            "Korean": "K",
            "Japanese": "J"
        }
        self.default_language = "E"  # Default to English

    def clean_numeric(self, value):
        """
        Clean numeric strings by removing commas and decimal points.
        
        Args:
            value: The value to clean.
        
        Returns:
            Cleaned numeric value.
        """
        if isinstance(value, str):
            return value.replace(',', '').split('.')[0]
        return value

    def round_to_nearest_30_seconds(self, seconds):
        """
        Round the given number of seconds to the nearest 30-second increment.
        
        Args:
            seconds: The number of seconds to round.
        
        Returns:
            Rounded number of seconds.
        """
        try:
            if pd.isna(seconds) or not str(seconds).strip():
                return 0
            return round(float(seconds) / 15) * 15
        except (ValueError, TypeError) as e:
            logging.warning(f"Error rounding seconds '{seconds}': {e}")
            return 0

    def load_and_clean_data(self, file_path: str) -> Optional[pd.DataFrame]:
        """
        Load data from the selected input file and perform initial transformations.
        
        Args:
            file_path: Path to the input file.
        
        Returns:
            Cleaned DataFrame or None if an error occurs.
        """
        try:
            logging.info(f"Loading data from {file_path}")
            
            # Read the main data
            df = pd.read_csv(file_path, skiprows=3)
            original_count = len(df)
            
            # Drop empty rows
            df = df.dropna(how='all')
            if len(df) < original_count:
                logging.warning(f"Dropped {original_count - len(df)} empty rows")
            
            # Filter required columns
            required_columns = ['id_contrattirighe', 'timerange2', 'dateschedule']
            before_required = len(df)
            df = df[df[required_columns].notna().all(axis=1)]
            if len(df) < before_required:
                logging.warning(f"Dropped {before_required - len(df)} rows with missing required values")
            
            # Filter out Textbox rows
            df = df[~df['IMPORTO2'].astype(str).str.contains('Textbox', na=False)]
            df = df[df.columns[~df.columns.str.contains('Textbox97|tot|Textbox61|Textbox53')]]
            
            # Clean numeric fields
            df['id_contrattirighe'] = df['id_contrattirighe'].apply(self.clean_numeric)
            if 'Textbox14' in df.columns:
                df['Textbox14'] = df['Textbox14'].apply(self.clean_numeric)
            
            # Rename columns
            column_mapping = {
                'id_contrattirighe': 'Line',
                'Textbox14': '#',
                'duration3': 'Length',
                'IMPORTO2': 'Gross Rate',
                'nome2': 'Market',
                'dateschedule': 'Air Date',
                'airtimep': 'Program',
                'bookingcode2': 'Media'
            }
            
            logging.info(f"Available columns before renaming: {df.columns.tolist()}")
            rename_dict = {k: v for k, v in column_mapping.items() if k in df.columns}
            df = df.rename(columns=rename_dict)
            logging.info(f"Columns after renaming: {df.columns.tolist()}")
            
            # Split timerange2
            if 'timerange2' in df.columns:
                df[['Time In', 'Time Out']] = df['timerange2'].str.split('-', expand=True)
            
            # Validate required columns after renaming
            df = df[df['Line'].notna()]
            
            if df.empty:
                raise ValueError("No valid data rows found after cleaning")
                
            return df
            
        except Exception as e:
            logging.error(f"Error in load_and_clean_data: {str(e)}")
            raise

    def apply_transformations(self, df: pd.DataFrame, text_box_180: str, text_box_171: str) -> pd.DataFrame:
        """
        Apply transformations including billcode generation and market replacements.
        
        Args:
            df: The DataFrame to transform.
            text_box_180: Value from Textbox180.
            text_box_171: Value from Textbox171.
        
        Returns:
            Transformed DataFrame.
        """
        try:
            # Set billcode
            billcode = f"{text_box_180}:{text_box_171}" if text_box_180 and text_box_171 else text_box_171 or text_box_180 or ''
            df['Bill Code'] = billcode
            
            # Check Market column
            if 'Market' not in df.columns:
                logging.error("Market column not found in DataFrame")
                logging.info(f"Available columns: {df.columns.tolist()}")
                raise KeyError("Market column not found in DataFrame")
            
            # Apply market replacements
            logging.info(f"Applying market replacements: {self.config.market_replacements}")
            df['Market'] = df['Market'].replace(self.config.market_replacements)
            
            # Transform Gross Rate
            if 'Gross Rate' in df.columns:
                df['Gross Rate'] = df['Gross Rate'].astype(str).str.replace('$', '').str.replace(',', '')
                df['Gross Rate'] = pd.to_numeric(df['Gross Rate'], errors='coerce').fillna(0).map("${:,.2f}".format)
            
            # Transform Length
            if 'Length' in df.columns:
                df['Length'] = df['Length'].apply(self.round_to_nearest_30_seconds)
                df['Length'] = pd.to_timedelta(df['Length'], unit='s').apply(lambda x: str(x).split()[-1].zfill(8))
            
            # Transform Line and #
            if 'Line' in df.columns:
                df['Line'] = pd.to_numeric(df['Line'], errors='coerce').fillna(0).astype(int)
            
            if '#' in df.columns:
                df['#'] = pd.to_numeric(df['#'], errors='coerce').fillna(0).astype(int)
            
            return df
            
        except Exception as e:
            logging.error(f"Error in transformations: {str(e)}")
            raise

    def detect_languages(self, df: pd.DataFrame) -> Tuple[Dict[str, int], pd.Series]:
        """
        Detect languages from the 'rowdescription' column.
        
        Args:
            df: The DataFrame to process.
        
        Returns:
            A tuple containing:
            - A dictionary of detected language counts.
            - A Series of language codes for each row.
        """
        languages = {}
        row_languages = pd.Series(index=df.index, dtype=str)
        
        if 'rowdescription' not in df.columns:
            logging.warning("No 'rowdescription' column found, defaulting to English")
            row_languages[:] = self.default_language
            languages[self.default_language] = len(df)
            return languages, row_languages
        
        # Process each row
        for idx, description in df['rowdescription'].items():
            if not isinstance(description, str):
                row_languages[idx] = self.default_language
                continue
            
            # Default to English until we find a match
            detected_lang = self.default_language
            
            # Try to match language names in order of longest to shortest
            sorted_mappings = sorted(
                [(k, v) for k, v in self.language_mapping.items() if k != 'default'],
                key=lambda x: len(x[0]),
                reverse=True
            )
            
            for keyword, code in sorted_mappings:
                if keyword.lower() in description.lower():
                    detected_lang = code
                    break
            
            row_languages[idx] = detected_lang
            languages[detected_lang] = languages.get(detected_lang, 0) + 1
        
        # Log detection results
        logging.info(f"Detected languages: {languages}")
        
        return languages, row_languages