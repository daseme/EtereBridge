# utils.py
import pandas as pd


def safe_convert_date(date_val: any) -> any:
    """Convert date_val to datetime; return None if unparseable."""
    try:
        return pd.to_datetime(date_val, errors="coerce")
    except Exception:
        return None
