# time_utils.py
import pandas as pd
import logging
from datetime import datetime
from typing import Optional

def unify_time_format(time_str: str, desired_format: str = "%H:%M:%S") -> Optional[str]:
    """
    Parse and convert time strings to a consistent format.
    Handles multiple input formats and safely returns None for invalid inputs.
    
    Args:
        time_str (str): Time string to parse
        desired_format (str): Format string for output time
        
    Returns:
        Optional[str]: Formatted time string or None if parsing fails
    """
    if not time_str or pd.isna(time_str):
        return None
        
    # Convert to string if needed
    if not isinstance(time_str, str):
        time_str = str(time_str)
    
    # Clean the string
    time_str = time_str.strip()
    
    # Try multiple formats in order of likelihood
    formats = ["%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M:%S %p"]
    
    for fmt in formats:
        try:
            parsed_time = datetime.strptime(time_str, fmt)
            return parsed_time.strftime(desired_format)
        except ValueError:
            continue
    
    # If all formats fail, log warning and return None
    logging.warning(f"Could not parse time string: '{time_str}'")
    return None

def transform_times(df: pd.DataFrame) -> pd.DataFrame:
    """
    Apply consistent time formatting to Time In and Time Out columns.
    
    Args:
        df (pd.DataFrame): DataFrame with time columns
        
    Returns:
        pd.DataFrame: DataFrame with transformed time columns
    """
    # Create a copy to avoid modifying the input DataFrame
    df = df.copy()
    
    for column in ["Time In", "Time Out"]:
        if column in df.columns:
            # Apply the unify_time_format function
            df[column] = df[column].apply(unify_time_format)
            
            # Replace None values with empty string to maintain compatibility
            df[column] = df[column].fillna("")
            
    return df

def excel_time_to_seconds(excel_time: float) -> int:
    """
    Convert Excel time value (fraction of day) to seconds.
    
    Args:
        excel_time (float): Excel time value (fraction of 24 hours)
        
    Returns:
        int: Time in seconds
    """
    if pd.isna(excel_time):
        return 0
    return int(excel_time * 24 * 60 * 60)

def seconds_to_excel_time(seconds: int) -> float:
    """
    Convert seconds to Excel time value (fraction of day).
    
    Args:
        seconds (int): Time in seconds
        
    Returns:
        float: Excel time value (fraction of 24 hours)
    """
    if pd.isna(seconds) or seconds == 0:
        return 0.0
    return seconds / (24 * 60 * 60)