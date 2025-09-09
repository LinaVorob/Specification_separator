from pathlib import Path
from typing import List

from logger import LoggerFile
from config.config import PROJECT_PATH, WORK_FORMATS, LOG_FILE_NAME, FINAL_FILE_NAME, NOT_NEED_COLUMNS
from excel import ExcelInput, ExcelOutput
from exceptions import IncorrectColumns


class Main:
    @staticmethod
    def list_excel_files() -> List[Path]:
        """List all Excel files (.xlsx, .xls) in the current directory.

        Scans the working directory for files with .xlsx or .xls extensions,
        displays enumerated results, and returns the list of matching files.

        Returns:
            list[str] | None: List of Excel file names if found, otherwise None

        Notes:
            - Prints "No Excel files found..." message if no files are present
            - Displays enumerated list of available files with indices
        """
        files = [f for f in PROJECT_PATH.iterdir() if f.suffix in WORK_FORMATS and not f.name.endswith(FINAL_FILE_NAME)]
        return files

    def main(self):
        input_handler = ExcelInput()
        output_handler = ExcelOutput()
        for excel_file in self.list_excel_files():
            logger = LoggerFile(log_file=LOG_FILE_NAME.format(excel_file.name.split('.')[0])).get_logger()
            try:
                input_handler.read_excel_file(excel_file)
                input_handler.work_with_rows()
                input_handler.delete_columns(NOT_NEED_COLUMNS)
                output_handler.relative_content = input_handler.get_data()
                output_handler.absolute_content = input_handler.models
                output_handler.write_excel_file(excel_file, input_handler.counter_unique_models)
            except IncorrectColumns:
                logger.error('Файл не соответствует формату. Пропуск.')
                continue

if __name__ == "__main__":
    Main().main()