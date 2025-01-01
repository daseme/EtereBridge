# config_setup.py
import logging
from pathlib import Path
from datetime import datetime

def setup_logging(output_dir: str) -> str:
    """Configure logging with both file and console output."""
    log_dir = Path(output_dir) / 'logs'
    log_dir.mkdir(parents=True, exist_ok=True)
    
    # Generate timestamp-based log filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f'processing_{timestamp}.log'
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    
    return str(log_file)