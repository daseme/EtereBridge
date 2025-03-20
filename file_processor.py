import pandas as pd
import re
from datetime import datetime
from typing import Dict, Tuple, Optional, Callable
import logging


# --- Pure Transformation Functions ---
def compute_broadcast_month(air_date: pd.Timestamp) -> pd.Timestamp:
    """
    Replicates your broadcast logic:
      - If the next Sunday crosses into next month (Dec->Jan), shift year.
      - Otherwise just use the month of that next Sunday, day=1.
    """
    if pd.isna(air_date):
        return None  # No date => no broadcast month
    # In pandas, weekday(): Monday=0, Sunday=6
    days_until_sunday = 6 - air_date.weekday()
    next_sunday = air_date + pd.Timedelta(days=days_until_sunday)

    # If we cross from Dec into Jan, increment year
    year = air_date.year
    if air_date.month == 12 and next_sunday.month == 1:
        year += 1

    return pd.Timestamp(year=year, month=next_sunday.month, day=1)


def transform_month_column(df: pd.DataFrame) -> pd.DataFrame:
    """
    Creates a real 'Month' column based on 'Air Date' and
    the value in 'Billing Type' (assumed to be 'Calendar' or 'Broadcast').

    If 'Billing Type' = 'Calendar', just use the Air Date.
    If 'Billing Type' = 'Broadcast', use the broadcast logic.
    """
    if "Air Date" not in df.columns:
        logging.warning("No 'Air Date' column found, cannot compute Month.")
        return df

    if "Billing Type" not in df.columns:
        logging.warning("No 'Billing Type' column found, defaulting all to Calendar.")
        df["Billing Type"] = "Calendar"

    def compute_month(row):
        ad = row["Air Date"]
        if pd.isna(ad):
            return None
        if row["Billing Type"] == "Calendar":
            return ad  # Just use the actual date
        else:
            return compute_broadcast_month(ad)

    # Convert Air Date to real datetime if possible
    df["Air Date"] = pd.to_datetime(df["Air Date"], errors="coerce")

    # Now create 'Month' column
    df["Month"] = df.apply(compute_month, axis=1)
    return df


def generate_billcode(first_part: str, second_part: str) -> str:
    """
    Generate bill code by combining two parts with a colon separator.
    
    The format is "first_part:second_part" when both parts exist.
    If only one part exists, return just that part without a colon.
    
    Args:
        first_part: The first part of the bill code (typically client/agency name)
        second_part: The second part of the bill code (typically site/venue name)
        
    Returns:
        Formatted bill code string
    """
    first_part = first_part.strip() if first_part else ""
    second_part = second_part.strip() if second_part else ""
    
    if first_part and second_part:
        return f"{first_part}:{second_part}"
    elif first_part:
        return first_part
    elif second_part:
        return second_part
    return ""


def apply_market_replacements(
    df: pd.DataFrame, market_replacements: Dict[str, str]
) -> pd.DataFrame:
    """Replace market names using provided mapping."""
    if "Market" not in df.columns:
        logging.error("Market column not found in DataFrame")
        logging.info(f"Available columns: {df.columns.tolist()}")
        raise KeyError("Market column not found in DataFrame")
    df["Market"] = df["Market"].replace(market_replacements)
    return df


def transform_gross_rate(
    df: pd.DataFrame, safe_to_numeric_func: Callable[[any], float]
) -> pd.DataFrame:
    """
    Clean and format the Gross Rate column by properly converting to numeric values.

    Args:
        df (pd.DataFrame): DataFrame with Gross Rate column to transform
        safe_to_numeric_func (Callable): Function to safely convert values to numeric type

    Returns:
        pd.DataFrame: DataFrame with transformed Gross Rate column
    """
    if "Gross Rate" in df.columns:
        # Convert string dollar values to actual numeric values
        df["Gross Rate"] = df["Gross Rate"].fillna(0)
        # Handle both string and numeric input by converting to string first
        df["Gross Rate"] = df["Gross Rate"].astype(str)
        # Remove currency symbols and thousands separators
        df["Gross Rate"] = df["Gross Rate"].str.replace("$", "", regex=False)
        df["Gross Rate"] = df["Gross Rate"].str.replace(",", "", regex=False)
        # Convert to numeric values for calculations (not strings)
        df["Gross Rate"] = pd.to_numeric(df["Gross Rate"], errors="coerce").fillna(0)
    return df


def transform_line_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Convert Line and '#' columns to integer."""
    if "Line" in df.columns:
        df["Line"] = pd.to_numeric(df["Line"], errors="coerce").fillna(0).astype(int)
    if "#" in df.columns:
        df["#"] = pd.to_numeric(df["#"], errors="coerce").fillna(0).astype(int)
    return df


# --- New Time Formatting Functions ---


def unify_time_format(time_str: str, desired_format: str = "%H:%M:%S") -> str:
    """
    Parse the time string and convert it to a consistent HH:MM:SS format.
    Returns the original value if parsing fails.
    """
    try:
        # Attempt parsing as H:M
        parsed = pd.to_datetime(time_str, format="%H:%M", errors="coerce")
        if pd.isna(parsed):
            # If parsing with %H:%M fails, try H:M:S
            parsed = pd.to_datetime(time_str, format="%H:%M:%S", errors="coerce")
        if pd.isna(parsed):
            # If neither format works, return original
            return time_str
        return parsed.strftime(desired_format)
    except Exception as e:
        logging.warning(f"Error unifying time format '{time_str}': {e}")
        return time_str


def transform_times(df: pd.DataFrame) -> pd.DataFrame:
    """
    Apply consistent time formatting to Time In and Time Out columns.
    """
    if "Time In" in df.columns:
        df["Time In"] = df["Time In"].astype(str).apply(unify_time_format)
    if "Time Out" in df.columns:
        df["Time Out"] = df["Time Out"].astype(str).apply(unify_time_format)
    return df


# --- FileProcessor Class ---


class FileProcessor:
    def __init__(self, config):
        """
        Initialize the FileProcessor with configuration settings.
        """
        self.config = config
        
        # Use language mapping from config instead of hardcoding
        self.language_mapping = {}
        if hasattr(config, 'config') and 'LanguageMapping' in config.config:
            self.language_mapping = dict(config.config['LanguageMapping'])
        
        # Get default language from config or use E as fallback
        self.default_language = self.language_mapping.get('default', 'E')

    def clean_numeric(self, value):
        """
        Clean numeric strings by removing commas and decimal parts.
        """
        if isinstance(value, str):
            return value.replace(",", "").split(".")[0]
        return value

    def round_to_nearest_increment(self, seconds):
        """
        Round the given seconds to the nearest 15-second increment.
        If the seconds are less than 15, return the original value.
        """
        try:
            if pd.isna(seconds) or not str(seconds).strip():
                return 0
            value = float(seconds)
            if value < 15:
                return value
            return round(value / 15) * 15
        except (ValueError, TypeError) as e:
            logging.warning(f"Error rounding seconds '{seconds}': {e}")
            return 0

    def transform_length(
        self, df: pd.DataFrame, round_func: Callable[[any], int]
    ) -> pd.DataFrame:
        """Transform the Length column by rounding and formatting."""
        if "Length" in df.columns:
            df["Length"] = df["Length"].apply(round_func)
            df["Length"] = (
                df["Length"].apply(self.round_to_nearest_increment).astype(int)
            )
        return df

    def safe_to_numeric(self, value):
        """
        Safely convert a value to numeric.
        """
        try:
            if pd.isna(value) or str(value).strip().lower() == "nan":
                return 0
            return pd.to_numeric(value, errors="raise")
        except ValueError as e:
            logging.warning(f"Failed to convert {value} to numeric: {e}")
            return 0

    def load_and_clean_data(self, file_path: str) -> Optional[pd.DataFrame]:
        """
        Load data from the selected input file and perform initial cleaning.
        Skips rows where 'dateschedule' is 'Unplaced', and prints the count.
        """
        try:
            logging.info(f"Loading data from {file_path}")
            df = pd.read_csv(file_path, skiprows=3)
            original_count = len(df)

            # Drop completely empty rows
            df = df.dropna(how="all")
            if len(df) < original_count:
                logging.warning(f"Dropped {original_count - len(df)} empty rows")

            # Check required columns
            required_columns = ["id_contrattirighe", "timerange2", "dateschedule"]
            before_required = len(df)
            df = df[df[required_columns].notna().all(axis=1)]
            if len(df) < before_required:
                logging.warning(
                    f"Dropped {before_required - len(df)} rows missing required columns"
                )

            # Skip rows containing "Textbox" in IMPORTO2
            df = df[~df["IMPORTO2"].astype(str).str.contains("Textbox", na=False)]

            # Drop columns that match certain patterns
            df = df[
                df.columns[
                    ~df.columns.str.contains("Textbox97|tot|Textbox61|Textbox53")
                ]
            ]

            # Skip rows where dateschedule == 'Unplaced'
            unplaced_count = df[
                df["dateschedule"].astype(str).str.lower() == "unplaced"
            ].shape[0]
            if unplaced_count > 0:
                print(
                    f"Skipping {unplaced_count} lines with 'Unplaced' in 'dateschedule'"
                )
                df = df[df["dateschedule"].astype(str).str.lower() != "unplaced"]

            # Clean numeric fields
            df["id_contrattirighe"] = df["id_contrattirighe"].apply(self.clean_numeric)
            if "Textbox14" in df.columns:
                df["Textbox14"] = df["Textbox14"].apply(self.clean_numeric)

            # Rename columns
            column_mapping = {
                "id_contrattirighe": "Line",
                "Textbox14": "#",
                "duration3": "Length",
                "IMPORTO2": "Gross Rate",
                "nome2": "Market",
                "dateschedule": "Air Date",
                "airtimep": "Program",
                "bookingcode2": "Media",
            }
            logging.info(f"Available columns before renaming: {df.columns.tolist()}")
            rename_dict = {k: v for k, v in column_mapping.items() if k in df.columns}
            df = df.rename(columns=rename_dict)
            logging.info(f"Columns after renaming: {df.columns.tolist()}")

            # Split timerange2 into Time In / Time Out
            if "timerange2" in df.columns:
                df[["Time In", "Time Out"]] = df["timerange2"].str.split(
                    "-", expand=True
                )

            # Ensure no empty "Line" entries
            df = df[df["Line"].notna()]
            if df.empty:
                raise ValueError("No valid data rows found after cleaning")

            return df

        except Exception as e:
            logging.error(f"Error in load_and_clean_data: {str(e)}")
            raise

    def apply_transformations(
        self, df: pd.DataFrame, first_part: str, second_part: str
    ) -> pd.DataFrame:
        """
        Apply data transformations by:
        - Generating the bill code from the two extracted parts.
        - Unifying Time In/Out formats.
        - Applying market replacements.
        - Transforming Gross Rate, Length, and line columns.
        """
        try:
            # Generate Bill Code
            from file_processor import generate_billcode  # Import the updated function
            billcode = generate_billcode(first_part, second_part)
            df["Bill Code"] = billcode
            logging.info(f"Generated Bill Code: {billcode}")

            # Unify Time In/Time Out to HH:MM:SS
            df = transform_times(df)

            # Apply market replacements
            df = apply_market_replacements(df, self.config.market_replacements)

            # Transform Gross Rate
            df = transform_gross_rate(df, self.safe_to_numeric)

            # Transform Length
            df = self.transform_length(df, self.round_to_nearest_increment)

            # Transform Line and '#' columns
            df = transform_line_columns(df)

            return df
        except Exception as e:
            logging.error(f"Error in transformations: {str(e)}")
            raise

    def detect_languages(self, df: pd.DataFrame) -> Tuple[Dict[str, int], pd.Series]:
        """
        Detect languages from the 'rowdescription' column with a weighted approach.
        Pre-compiles regex patterns once for performance and uses reliable mapping.
        
        Args:
            df (pd.DataFrame): DataFrame containing rowdescription column
            
        Returns:
            Tuple[Dict[str, int], pd.Series]: 
                - Dictionary of language counts
                - Series mapping row indices to detected language codes
        """
        # Initialize return values
        languages = {}
        row_languages = pd.Series(index=df.index, dtype=str)
        
        # Check if required column exists
        if "rowdescription" not in df.columns:
            logging.warning("No 'rowdescription' column found, defaulting to English")
            default_lang = self.default_language
            row_languages[:] = default_lang
            languages[default_lang] = len(df)
            return languages, row_languages
        
        # Get program mappings from config
        program_language_map = self.config.program_language_map or {}
        default_language = getattr(self, 'default_language', 'E')
        
        # Compile regex patterns only once when the class is initialized
        if not hasattr(self, '_compiled_patterns'):
            self._compiled_patterns = {}
            pattern_mappings = {
                r'\bviet\b': 'V', 
                r'\bvietnamese\b': 'V',
                r'\bchinese\b': 'M',
                r'\bmandarin\b': 'M',
                r'\bcantonese\b': 'C',
                r'\bfilipino\b': 'T',
                r'\btagalog\b': 'T',
                r'\bhmong\b': 'Hm',
                r'\bkorean\b': 'K',
                r'\bjapanese\b': 'J',
                r'\bsouth asian\b': 'SA',
                r'\bhindi\b': 'SA',
                r'\bpunjabi\b': 'SA',
            }
            for term, lang_code in pattern_mappings.items():
                self._compiled_patterns[re.compile(term, re.IGNORECASE)] = lang_code

        # Process each row for language detection
        for idx, description in df["rowdescription"].items():
            # Handle non-string values
            if not isinstance(description, str):
                row_languages[idx] = default_language
                languages[default_language] = languages.get(default_language, 0) + 1
                continue
                
            # Initialize language scores with default language having a small baseline
            language_scores = {lang: 0 for lang in self.config.language_options}
            language_scores[default_language] = 1  # Default language gets a small baseline score
            
            # 1. Check exact program name matches (highest weight)
            description_lower = description.lower()  # Convert once for all comparisons
            for program, lang in program_language_map.items():
                if program.lower() in description_lower:
                    language_scores[lang] = language_scores.get(lang, 0) + 10
            
            # 2. Check language keyword matches (medium weight)
            for keyword, lang in self.language_mapping.items():
                if keyword.lower() in description_lower:
                    language_scores[lang] = language_scores.get(lang, 0) + 5
            
            # 3. Check regex pattern matches (lower weight)
            for pattern, lang in self._compiled_patterns.items():
                if pattern.search(description):
                    language_scores[lang] = language_scores.get(lang, 0) + 3
            
            # Determine the best language match
            best_lang, _ = max(language_scores.items(), key=lambda x: x[1])
            
            # Record the language for this row
            row_languages[idx] = best_lang
            languages[best_lang] = languages.get(best_lang, 0) + 1
        
        logging.info(f"Detected languages: {languages}")
        return languages, row_languages