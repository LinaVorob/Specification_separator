"""
Excel File Handler Module

Provides functionality for listing and interacting with Excel files in the current directory.
"""
from pathlib import Path

import pandas as pd
import os

from config.config import NECESSARY_COLUMNS, FINAL_FILE_NAME, LOG_FILE_NAME
from exceptions import IncorrectColumns, IncorrectRow
from logger import LoggerFile


class ExcelOutput:
    """Class for handling Excel file operations in the current working directory."""

    def __init__(self):
        self.relative_content: pd.DataFrame = None
        self.absolute_content: pd.DataFrame = None
        self.logger = None

    def write_excel_file(self, file_name: Path, file_contents: pd.DataFrame) -> None:
        self.logger = LoggerFile(name=LOG_FILE_NAME.format(file_name.name.split('.')[0])).get_logger()
        name = FINAL_FILE_NAME.format(file_name.name.rsplit('.', 1)[0])
        with pd.ExcelWriter(name, mode="w", engine="openpyxl") as writer:
            file_contents.to_excel(writer)

    def create_sheet(self, sheet_name: str):
        ...

    def create_absolute_sheet(self):
        ...

    def create_relative_sheet(self):
        ...


class ExcelInput:
    def __init__(self):
        self.file_contents = None
        self.fixed_content = None
        self.logger = None

    def read_excel_file(self, file_name: Path) -> pd.DataFrame:
        engine = 'xlrd'
        if file_name.suffix == ".xlsx":
            engine = 'openpyxl'
        self.logger = LoggerFile(name=LOG_FILE_NAME.format(file_name.name.split('.')[0])).get_logger()
        self.file_contents = pd.read_excel(str(file_name), engine=engine, na_values=[' ', '\t', '\n', '', '-'])
        self.fixed_content = self.file_contents.copy()
        self.delete_word_break()
        self.check_columns_kit()

    def delete_word_break(self) -> None:
        for column in self.file_contents.columns:
            if '\n' in column:
                correct_name = column.replace('\n', '')
                self.file_contents = self.file_contents.rename(columns={column: correct_name})

    def get_count_of_parts(self):
        ...

    def fix_origin_file(self) -> None:
        for index, row in self.file_contents.iterrows():
            try:
                self.check_row_number(index, row)
                self.check_word_break(index, row)
                self.check_empty_row(index, row)
                self.check_title_column(index, row)
            except IncorrectRow as err:
                self.logger.error(str(err))
                continue

    def check_row_number(self, index: int, row: pd.Series):
        self.file_contents.iloc[:, 0] = self.file_contents.iloc[:, 0].astype(str)
        for index, row in self.file_contents.iterrows():
            row.iloc[0] = row.iloc[0].replace(',', '.')

    def check_word_break(self, index: int, row: pd.Series):
        ...

    def check_empty_row(self, index: int, row: pd.Series):
        for index, row in self.file_contents.iterrows():
            if row.iloc[1:-1].isna().all():
                self.fixed_content = self.file_contents.drop(row.index)

    def check_title_column(self, index: int, row: pd.Series):
        ...

    def check_columns_kit(self):
        if not all([column in self.file_contents.columns for column in NECESSARY_COLUMNS]):
            raise IncorrectColumns

    def get_data(self) -> pd.DataFrame:
        return self.file_contents
