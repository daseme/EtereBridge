# EtereBridge

EtereBridge is a robust Python tool designed to process and transform Excel and CSV files for media planning and advertising workflows. It provides a user-friendly interface for handling complex data transformations, including multi-language support, market replacements, and specialized handling for WorldLink orders.

## Features

- **Interactive Processing**: Choose between batch processing or file-by-file handling
- **Flexible Input Handling**: Supports CSV files with various formats and configurations
- **Language Detection**: Automatic detection and verification of multiple languages in content
- **Market Replacements**: Configurable market name standardization
- **WorldLink Order Support**: Special handling for WorldLink orders with automated makegood processing
- **Comprehensive Logging**: Detailed logging system for tracking processing steps and errors
- **Progress Tracking**: Visual progress indicators for batch processing
- **Error Recovery**: Robust error handling with detailed feedback
- **Interim Results**: Automatic saving of interim results to prevent data loss

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/eterebridge.git
cd eterebridge
```

2. Create a virtual environment and activate it:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Configuration

The tool uses a `config.ini` file for configuration. Create this file in the root directory with the following sections:

```ini
[Paths]
template_path = templates/template.xlsx
input_dir = input
output_dir = output

[Sales]
sales_people = Person1,Person2,Person3

[Markets]
# Add market replacements here
Market1 = StandardName1
Market2 = StandardName2

[Columns]
final_columns = Column1,Column2,Column3

[Languages]
options = E,M,T,Hm,SA,V,C,K,J

[Type]
options = COM,PSA,PGM
```

## Usage

1. Place your input CSV files in the configured input directory.

2. Run the main script:
```bash
python EtereBridge.py
```

3. Follow the interactive prompts to:
   - Choose between batch or individual file processing
   - Select WorldLink or standard processing
   - Provide necessary input parameters (billing type, revenue type, etc.)
   - Verify language detection
   - Configure additional processing options

4. Check the output directory for processed files and logs.

## File Structure

```
eterebridge/
├── EtereBridge.py        # Main application file
├── config_manager.py     # Configuration management
├── file_processor.py     # File processing logic
├── config.ini           # Configuration file
├── input/              # Input directory for CSV files
├── output/             # Output directory for processed files
└── templates/          # Excel templates
```

## Processing Steps

1. **File Loading**: The tool loads CSV files and performs initial validation
2. **Data Cleaning**: Removes empty rows and unnecessary columns
3. **Language Detection**: Automatically detects languages in content
4. **Transformations**: Applies configured transformations including:
   - Market name standardization
   - Numeric value formatting
   - Time and date formatting
   - Billcode generation
5. **User Input Collection**: Gathers necessary processing parameters
6. **Output Generation**: Creates formatted Excel files with proper formatting and formulas

## Error Handling

- Detailed error messages for common issues
- Automatic logging of all processing steps
- Interim result saving to prevent data loss
- User-friendly error reporting in the console

## Troubleshooting

Common issues and solutions:

1. **Missing Configuration**:
   - Ensure `config.ini` exists and contains all required sections
   - Check file paths in configuration are correct

2. **Input File Issues**:
   - Verify CSV file format matches expected structure
   - Check for special characters in headers
   - Ensure required columns are present

3. **Output Errors**:
   - Verify write permissions in output directory
   - Check template file exists and is accessible

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## License

[Specify your license here]

## Support

For support and questions:
- Create an issue in the repository
- Contact [your contact information]