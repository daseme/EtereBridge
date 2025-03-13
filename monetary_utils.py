# monetary_utils.py
import pandas as pd
import logging
import re
from typing import List, Optional

def standardize_monetary_columns(df: pd.DataFrame, monetary_columns=None) -> pd.DataFrame:
    """
    Standardize all monetary columns by properly converting to numeric values.
    Handles various input formats including currency symbols, dashes, and empty values.
    
    Args:
        df (pd.DataFrame): DataFrame with monetary columns to transform
        monetary_columns (List[str], optional): List of column names to process.
            If None, defaults to ["Gross Rate", "Spot Value", "Station Net", "Broker Fees"]
        
    Returns:
        pd.DataFrame: DataFrame with transformed monetary columns
    """
    # Default monetary columns if not specified
    if monetary_columns is None:
        monetary_columns = ["Gross Rate", "Spot Value", "Station Net", "Broker Fees"]
    
    # Create a copy to avoid modifying the input DataFrame
    df = df.copy()
    
    # Filter to only include columns that actually exist in the DataFrame
    cols_to_process = [col for col in monetary_columns if col in df.columns]
    
    if not cols_to_process:
        logging.warning("No monetary columns found in DataFrame")
        return df
    
    # Define a function to clean currency values
    def clean_currency(x):
        if pd.isna(x) or x == '' or (isinstance(x, str) and x.strip() in ['-', 'N/A']):
            return 0
        
        # If it's already a number, return it
        if isinstance(x, (int, float)):
            return float(x)
            
        # Otherwise convert from string
        # Remove currency symbols, commas and whitespace
        cleaned = re.sub(r'[$,\s]', '', str(x))
        try:
            return float(cleaned)
        except (ValueError, TypeError):
            logging.warning(f"Could not convert value '{x}' to number, using 0 instead")
            return 0
    
    # Process each monetary column
    for col in cols_to_process:
        # First handle blank/null values - standardize to 0
        df[col] = df[col].fillna(0)
        
        # Apply the cleaning function
        df[col] = df[col].apply(clean_currency)
        
        # Ensure the column is numeric type with proper formatting
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        
        logging.info(f"Standardized monetary column: {col}")
    
    return df

def format_excel_monetary_columns(sheet, df, monetary_columns=None, start_row=2):
    """
    Apply consistent currency formatting to monetary columns in an Excel worksheet.
    
    Args:
        sheet: The Excel worksheet object
        df (pd.DataFrame): The DataFrame being written to Excel
        monetary_columns (List[str], optional): List of monetary column names.
            If None, defaults to ["Gross Rate", "Spot Value", "Station Net", "Broker Fees"]
        start_row (int): The starting row number in the Excel sheet (typically 2, after headers)
    """
    # Default monetary columns if not specified
    if monetary_columns is None:
        monetary_columns = ["Gross Rate", "Spot Value", "Station Net", "Broker Fees"]
    
    # Get column indices for monetary columns
    monetary_indices = {}
    for col_name in monetary_columns:
        if col_name in df.columns:
            col_idx = df.columns.get_loc(col_name) + 1  # +1 because Excel is 1-indexed
            monetary_indices[col_name] = col_idx
    
    if not monetary_indices:
        logging.warning("No monetary columns found for Excel formatting")
        return
    
    # Define currency format string
    currency_format = '"$"#,##0.00_);("$"#,##0.00)'
    
    # Apply formatting to all cells in monetary columns
    for col_name, col_idx in monetary_indices.items():
        for row_num in range(start_row, len(df) + start_row):
            cell = sheet.cell(row=row_num, column=col_idx)
            
            # Ensure cell has a value (even if zero)
            if cell.value is None or cell.value == '':
                cell.value = 0
                
            # Apply currency format
            cell.number_format = currency_format
            
        logging.info(f"Applied currency formatting to Excel column: {col_name}")