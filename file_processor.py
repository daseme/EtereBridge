import pandas as pd
from typing import Dict, Tuple, Optional, Callable
import logging

# --- Pure Transformation Functions ---

def generate_billcode(text_box_180: str, text_box_171: str) -> str:
    """Generate bill code by combining two text boxes."""
    if text_box_180 and text_box_171:
        return f"{text_box_180}:{text_box_171}"
    elif text_box_171:
        return text_box_171
    elif text_box_180:
        return text_box_180
    return ""

def apply_market_replacements(df: pd.DataFrame, market_replacements: Dict[str, str]) -> pd.DataFrame:
    """Replace market names using provided mapping."""
    if "Market" not in df.columns:
        logging.error("Market column not found in DataFrame")
        logging.info(f"Available columns: {df.columns.tolist()}")
        raise KeyError("Market column not found in DataFrame")
    df["Market"] = df["Market"].replace(market_replacements)
    return df

def transform_gross_rate(df: pd.DataFrame, safe_to_numeric_func: Callable[[any], float]) -> pd.DataFrame:
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

def transform_length(df: pd.DataFrame, round_func: Callable[[any], int]) -> pd.DataFrame:
    """Transform the Length column by rounding and formatting."""
    if "Length" in df.columns:
        df["Length"] = df["Length"].apply(round_func)
        df["Length"] = pd.to_timedelta(df["Length"], unit="s").apply(lambda x: str(x).split()[-1].zfill(8))
    return df

def transform_line_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Convert Line and '#' columns to integer."""
    if "Line" in df.columns:
        df["Line"] = pd.to_numeric(df["Line"], errors="coerce").fillna(0).astype(int)
    if "#" in df.columns:
        df["#"] = pd.to_numeric(df["#"], errors="coerce").fillna(0).astype(int)
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

    def round_to_nearest_30_seconds(self, seconds):
        """
        Round the given seconds to the nearest 30-second increment.
        """
        try:
            if pd.isna(seconds) or not str(seconds).strip():
                return 0
            return round(float(seconds) / 15) * 15
        except (ValueError, TypeError) as e:
            logging.warning(f"Error rounding seconds '{seconds}': {e}")
            return 0

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
        """
        try:
            logging.info(f"Loading data from {file_path}")
            df = pd.read_csv(file_path, skiprows=3)
            original_count = len(df)
            df = df.dropna(how="all")
            if len(df) < original_count:
                logging.warning(f"Dropped {original_count - len(df)} empty rows")
            required_columns = ["id_contrattirighe", "timerange2", "dateschedule"]
            before_required = len(df)
            df = df[df[required_columns].notna().all(axis=1)]
            if len(df) < before_required:
                logging.warning(f"Dropped {before_required - len(df)} rows with missing required values")
            df = df[~df["IMPORTO2"].astype(str).str.contains("Textbox", na=False)]
            df = df[df.columns[~df.columns.str.contains("Textbox97|tot|Textbox61|Textbox53")]]
            df["id_contrattirighe"] = df["id_contrattirighe"].apply(self.clean_numeric)
            if "Textbox14" in df.columns:
                df["Textbox14"] = df["Textbox14"].apply(self.clean_numeric)
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
            if "timerange2" in df.columns:
                df[["Time In", "Time Out"]] = df["timerange2"].str.split("-", expand=True)
            df = df[df["Line"].notna()]
            if df.empty:
                raise ValueError("No valid data rows found after cleaning")
            return df
        except Exception as e:
            logging.error(f"Error in load_and_clean_data: {str(e)}")
            raise

    def apply_transformations(self, df: pd.DataFrame, text_box_180: str, text_box_171: str) -> pd.DataFrame:
        """
        Apply data transformations by:
         - Generating the bill code.
         - Applying market replacements.
         - Transforming Gross Rate, Length, and line columns.
        """
        try:
            # Generate bill code and assign it to 'Bill Code'
            billcode = generate_billcode(text_box_180, text_box_171)
            df["Bill Code"] = billcode

            # Apply market replacements
            df = apply_market_replacements(df, self.config.market_replacements)

            # Transform Gross Rate
            df = transform_gross_rate(df, self.safe_to_numeric)

            # Transform Length
            df = transform_length(df, self.round_to_nearest_30_seconds)

            # Transform Line and '#' columns
            df = transform_line_columns(df)

            return df
        except Exception as e:
            logging.error(f"Error in transformations: {str(e)}")
            raise

    def detect_languages(self, df: pd.DataFrame) -> Tuple[Dict[str, int], pd.Series]:
        """
        Detect languages from the 'rowdescription' column.
        """
        languages = {}
        row_languages = pd.Series(index=df.index, dtype=str)
        if "rowdescription" not in df.columns:
            logging.warning("No 'rowdescription' column found, defaulting to English")
            row_languages[:] = self.default_language
            languages[self.default_language] = len(df)
            return languages, row_languages
        for idx, description in df["rowdescription"].items():
            if not isinstance(description, str):
                row_languages[idx] = self.default_language
                continue
            detected_lang = self.default_language
            sorted_mappings = sorted(
                [(k, v) for k, v in self.language_mapping.items() if k != "default"],
                key=lambda x: len(x[0]),
                reverse=True,
            )
            for keyword, code in sorted_mappings:
                if keyword.lower() in description.lower():
                    detected_lang = code
                    break
            row_languages[idx] = detected_lang
            languages[detected_lang] = languages.get(detected_lang, 0) + 1
        logging.info(f"Detected languages: {languages}")
        return languages, row_languages
