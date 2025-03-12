import pandas as pd
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


def generate_billcode(text_box_180: str, text_box_171: str) -> str:
    """Generate bill code by combining two text boxes."""
    if text_box_180 and text_box_171:
        return f"{text_box_180}:{text_box_171}"
    elif text_box_171:
        return text_box_171
    elif text_box_180:
        return text_box_180
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
    """Clean and format the Gross Rate column."""
    if "Gross Rate" in df.columns:
        df["Gross Rate"] = (
            df["Gross Rate"]
            .fillna(0)
            .astype(str)
            .str.strip()
            .str.replace("$", "")
            .str.replace(",", "")
        )
        df["Gross Rate"] = df["Gross Rate"].apply(safe_to_numeric_func).fillna(0)
        df["Gross Rate"] = df["Gross Rate"].map("${:,.2f}".format)
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
        self.language_mapping = {
            "Chinese": "M",
            "Filipino": "T",
            "Hmong": "Hm",
            "South Asian": "SA",
            "Vietnamese": "V",
            "Mandarin": "M",
            "Cantonese": "C",
            "Korean": "K",
            "Japanese": "J",
        }
        self.default_language = "E"  # Default to English

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
        
    def transform_length(self,
    df: pd.DataFrame, round_func: Callable[[any], int]
) -> pd.DataFrame:
        """Transform the Length column by rounding and formatting."""
        if "Length" in df.columns:
            df["Length"] = df["Length"].apply(round_func)
            df["Length"] = df["Length"].apply(self.round_to_nearest_increment).astype(int)        
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
        self, df: pd.DataFrame, text_box_180: str, text_box_171: str
    ) -> pd.DataFrame:
        """
        Apply data transformations by:
         - Generating the bill code.
         - Unifying Time In/Out formats.
         - Applying market replacements.
         - Transforming Gross Rate, Length, and line columns.
        """
        try:
            # Bill code
            billcode = generate_billcode(text_box_180, text_box_171)
            df["Bill Code"] = billcode

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
        Detect languages from the 'rowdescription' column.
        Uses both keyword matching and pattern recognition for better accuracy.
        """
        languages = {}
        row_languages = pd.Series(index=df.index, dtype=str)

        if "rowdescription" not in df.columns:
            logging.warning("No 'rowdescription' column found, defaulting to English")
            row_languages[:] = self.default_language
            languages[self.default_language] = len(df)
            return languages, row_languages

        # Define additional language patterns with regexes
        # \b means word boundary - matches spaces, punctuation, etc.
        language_patterns = {
            r'\bviet\b': 'V',         # Matches "viet" as a stand-alone word
            r'\bvietnamese\b': 'V',   # Matches "vietnamese" as a stand-alone word
            r'\bchinese\b': 'M',      # Matches "chinese" as a stand-alone word
            r'\bfilipino\b': 'T',     # Matches "filipino" as a stand-alone word
            r'\btagalog\b': 'T',      # Matches "tagalog" as a stand-alone word  
            r'\bhmong\b': 'Hm',       # Matches "hmong" as a stand-alone word
            r'\bkorean\b': 'K',       # Matches "korean" as a stand-alone word
            r'\bjapanese\b': 'J',     # Matches "japanese" as a stand-alone word
            r'\bsouth asian\b': 'SA', # Matches "south asian" as a stand-alone phrase
        }

        for idx, description in df["rowdescription"].items():
            if not isinstance(description, str):
                row_languages[idx] = self.default_language
                continue

            # Default to English unless we find a match
            detected_lang = self.default_language
            
            # Convert to lowercase for case-insensitive matching
            desc_lower = description.lower()
            
            # Check keyword dictionary first (existing method)
            sorted_mappings = sorted(
                [(k, v) for k, v in self.language_mapping.items() if k != "default"],
                key=lambda x: len(x[0]),
                reverse=True,
            )
            for keyword, code in sorted_mappings:
                if keyword.lower() in desc_lower:
                    detected_lang = code
                    break
            
            # If still English, try pattern matching
            if detected_lang == self.default_language:
                import re
                for pattern, code in language_patterns.items():
                    if re.search(pattern, desc_lower):
                        detected_lang = code
                        break

            row_languages[idx] = detected_lang
            languages[detected_lang] = languages.get(detected_lang, 0) + 1

        logging.info(f"Detected languages: {languages}")
        return languages, row_languages
