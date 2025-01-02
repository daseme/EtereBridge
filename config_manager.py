import os
import sys
import configparser
from dataclasses import dataclass
from typing import Dict, List
from config_setup import setup_logging


@dataclass
class Paths:
    """Container for application paths."""

    template_path: str
    input_dir: str
    output_dir: str
    script_dir: str


@dataclass
class AppConfig:
    """Container for application configuration."""

    paths: Paths
    market_replacements: Dict[str, str]
    final_columns: List[str]
    sales_people: List[str]
    language_options: List[str]
    type_options: List[str]


class ConfigurationError(Exception):
    """Custom exception for configuration-related errors."""

    pass


class ConfigManager:
    """Manages application configuration loading and validation."""

    def __init__(self, config_file: str = "config.ini"):
        """Initialize the configuration manager.

        Args:
            config_file (str): Name of the configuration file (default: 'config.ini')
        """
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_path = os.path.join(self.script_dir, config_file)
        self.config = self._load_config_file()
        self.app_config = self._create_app_config()

    def setup_logging(self) -> str:
        """Set up logging using the configured output directory."""
        return setup_logging(self.app_config.paths.output_dir)

    def _load_config_file(self) -> configparser.ConfigParser:
        """Load and validate the configuration file.

        Returns:
            configparser.ConfigParser: Loaded configuration

        Raises:
            ConfigurationError: If config file is missing or invalid
        """
        config = configparser.ConfigParser()
        config.optionxform = str  # Preserve case sensitivity of keys

        if not os.path.exists(self.config_path):
            raise ConfigurationError(f"Config file not found at: {self.config_path}")

        files_read = config.read(self.config_path)
        if not files_read:
            raise ConfigurationError(
                f"Could not read config file at: {self.config_path}"
            )

        return config

    def _validate_required_sections(self):
        """Validate that all required sections and keys exist in config.

        Raises:
            ConfigurationError: If required sections or keys are missing
        """
        required_sections = {
            "Paths": ["template_path", "input_dir", "output_dir"],
            "Sales": ["sales_people"],
            "Markets": [],  # Empty list means we just check for section existence
            "Columns": ["final_columns"],
        }

        for section, keys in required_sections.items():
            if section not in self.config:
                raise ConfigurationError(f"Missing required section: [{section}]")

            for key in keys:
                if key not in self.config[section]:
                    raise ConfigurationError(
                        f"Missing required key '{key}' in section [{section}]"
                    )

    def _create_paths(self) -> Paths:
        """Create and validate application paths.

        Returns:
            Paths: Container with validated application paths
        """
        paths = Paths(
            template_path=os.path.join(
                self.script_dir, self.config["Paths"]["template_path"]
            ),
            input_dir=os.path.join(self.script_dir, self.config["Paths"]["input_dir"]),
            output_dir=os.path.join(
                self.script_dir, self.config["Paths"]["output_dir"]
            ),
            script_dir=self.script_dir,
        )

        # Create directories if they don't exist
        os.makedirs(paths.output_dir, exist_ok=True)
        os.makedirs(os.path.dirname(paths.template_path), exist_ok=True)
        os.makedirs(paths.input_dir, exist_ok=True)

        return paths

    def _create_app_config(self) -> AppConfig:
        """Create the application configuration container.

        Returns:
            AppConfig: Container with validated configuration

        Raises:
            ConfigurationError: If configuration is invalid
        """
        self._validate_required_sections()

        # Create paths
        paths = self._create_paths()

        # Load market replacements
        market_replacements = dict(self.config["Markets"])

        # Load and parse final columns
        final_columns = self.config["Columns"]["final_columns"].split(",")
        final_columns = [
            col.strip() if col.strip() != "Number" else "#" for col in final_columns
        ]

        # Load sales people
        sales_people = self.config["Sales"]["sales_people"].split(",")

        # Load language options
        language_options = [
            opt.strip() for opt in self.config["Languages"]["options"].split(",")
        ]

        # Load type options
        type_options = [
            opt.strip() for opt in self.config["Type"]["options"].split(",")
        ]

        return AppConfig(
            paths=paths,
            market_replacements=market_replacements,
            final_columns=final_columns,
            sales_people=sales_people,
            language_options=language_options,
            type_options=type_options,
        )

    def get_config(self) -> AppConfig:
        """Get the application configuration.

        Returns:
            AppConfig: The validated application configuration
        """
        return self.app_config


# Create a global instance
try:
    config_manager = ConfigManager()
except ConfigurationError as e:
    print(f"Configuration error: {e}")
    print("Please ensure your config.ini file is properly set up.")
    sys.exit(1)
except Exception as e:
    print(f"Unexpected error loading configuration: {e}")
    sys.exit(1)
