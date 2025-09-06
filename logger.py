import logging
from datetime import datetime
from pathlib import Path

from config.config import PROJECT_PATH


class LoggerFile:
    """Class for setting up and managing logging to both console and file."""
    _instance = None  # Class-level variable to store the singleton instance

    def __new__(cls, *args, **kwargs):
        """Ensure only one instance of LoggerFile is created (Singleton pattern)."""
        if cls._instance is None:
            cls._instance = super(LoggerFile, cls).__new__(cls)
        return cls._instance

    def __init__(self, name: str = __name__, log_file: str = "app.txt",
                 console_level=logging.DEBUG, file_level=logging.INFO):
        """
        Initialize the logger with a name, log file path, and logging level.

        Args:
            name (str): Name of the logger.
            log_file (str): Path to the log file.
            level (int): Logging level (e.g., logging.DEBUG, logging.INFO).
        """
        self.name: str = name
        self.log_file: Path = PROJECT_PATH / log_file
        self.console_level: int = console_level
        self.file_level: int = file_level
        self.logger: logging.Logger = logging.getLogger(name)
        self.check_file_exists()
        self._setup_logger()

    def _setup_logger(self):
        """Configure the logger to write to console and file with the specified format."""
        self.logger.setLevel(logging.DEBUG)

        # Create a formatter for the log messages
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )

        # File handler to write logs to a file
        file_handler = logging.FileHandler(self.log_file)
        file_handler.setLevel(self.file_level)
        file_handler.setFormatter(formatter)

        # Console handler to output logs to the terminal
        console_handler = logging.StreamHandler()
        console_handler.setLevel(self.console_level)
        console_handler.setFormatter(formatter)

        # Add handlers to the logger
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)

    def get_logger(self):
        """Return the configured logger instance."""
        return self.logger

    def check_file_exists(self):
        if not self.log_file.parent.exists():
            for path in self.log_file.parents:
                if path.exists():
                    break
                path.mkdir(parents=True)
        if not self.log_file.exists():
            self.log_file.touch()


# Example usage
if __name__ == "__main__":
    logger = LoggerFile(log_file=f"logs/app_{datetime.now().strftime('%Y%m%d')}.txt").get_logger()
    logger.debug("This is a debug message.")
    logger.info("This is an info message.")
    logger.warning("This is a warning message.")
    logger.error("This is an error message.")
    logger.critical("This is a critical message.")