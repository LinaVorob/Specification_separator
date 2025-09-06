"""
Excel File Handler Module

Provides functionality for listing and interacting with Excel files in the current directory.
"""
import logging
from pathlib import Path
from typing import Tuple

import numpy as np
import pandas as pd

from config.config import NECESSARY_COLUMNS, FINAL_FILE_NAME, LOG_FILE_NAME, RELATIVE_SHEET_NAME
from exceptions import IncorrectColumns, IncorrectRow
from logger import LoggerFile
from models import DetailTypes


class ExcelOutput:
    """Class for handling Excel file operations in the current working directory."""

    def __init__(self):
        self.relative_content: pd.DataFrame = None
        self.absolute_content: pd.DataFrame = None
        self.logger = None

    def write_excel_file(self, file_name: Path, file_contents: pd.DataFrame) -> None:
        self.logger = LoggerFile(name=LOG_FILE_NAME.format(file_name.name.split('.')[0])).get_logger()
        name = f'{file_name.name.rsplit('.', 1)[0]}{FINAL_FILE_NAME}'
        with pd.ExcelWriter(name, mode="w", engine="openpyxl") as writer:
            file_contents.to_excel(writer, sheet_name=RELATIVE_SHEET_NAME, index=False)

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
        """
        Read Excel file
        :param file_name: file name
        :return: DataFrame with file contents
        """
        engine = 'xlrd'
        if file_name.suffix == ".xlsx":
            engine = 'openpyxl'
        self.logger = LoggerFile(name=LOG_FILE_NAME.format(file_name.name.split('.')[0]), file_level=logging.INFO).get_logger()
        self.logger.info(f'Read file {file_name.name}')
        self.file_contents = pd.read_excel(str(file_name), engine=engine, na_values=['  ', ' ', '\t', '\n', '', '-', '  ', '   '])
        self.replace_spaces_to_none()
        self.delete_word_break()
        self.check_columns_kit()

    def delete_word_break(self) -> None:
        """
        Delete word breaks from headers
        """
        self.logger.debug('Delete word breaks from headers')
        for column in self.file_contents.columns:
            if '\n' in column:
                correct_name = column.replace('\n', '')
                self.file_contents = self.file_contents.rename(columns={column: correct_name})

    def get_count_of_parts(self):
        ...

    def fix_origin_file(self) -> None:
        """
        Do all fixes for file
        """
        empty_index = list()
        self.file_contents.iloc[:, 0] = self.file_contents.iloc[:, 0].astype(str)

        for index, row in self.file_contents.iterrows():
            try:
                self.check_row_number(row)
                self.file_contents.iloc[index] = self.check_word_break(row)
                if self.check_empty_row(row):
                    empty_index.append(index)
                self.file_contents.iloc[index] = self.check_title_column(row)
            except IncorrectRow as err:
                self.logger.error(str(err))
                continue
        else:
            self.file_contents = self.file_contents.drop(empty_index).reset_index(drop=True)
            self.fix_row_number_sequence()

    def check_row_number(self, row: pd.Series) -> None:
        """
        Check a row number is not contains commas
        :param row: row of file
        """
        self.logger.debug('Check row number')
        row.iloc[0] = row.iloc[0].replace(',', '.')

    def check_word_break(self, row: pd.Series) -> pd.Series:
        """
        Check a row does not contain word breaks
        :param row: row of file
        """
        for index, cell in row.items():
            self.logger.debug(f'Check word break in cell {index}')
            if not pd.isna(cell):
                cell = str(cell)
                row.loc[index] = cell.replace('\n', ' ').replace('\n', ' ').replace('\t', ' ').replace('\r', ' ').replace('\xa0', ' ').replace('  ', ' ')
        return row

    def check_empty_row(self, row: pd.Series) -> bool:
        """
        Check a row is empty
        :param row: row of file
        :return: True if row is empty, False otherwise
        """
        self.logger.debug('Check empty row')
        return row.iloc[1:-1].isna().all()

    def check_title_column(self, row: pd.Series) -> pd.Series:
        """
        For all standard details check that title column is not empty
        :param row: row of file
        """
        if row.iloc[5] == DetailTypes.detail.value and row.iloc[1].empty:
            self.logger.debug('Check title column')
            row.iloc[1] = row.iloc[4]
        return row

    def check_columns_kit(self) -> None:
        """
        Check if file has all necessary columns
        """
        self.logger.debug('Check columns\' kit')
        if not all([column in self.file_contents.columns for column in NECESSARY_COLUMNS]):
            self.logger.error('File has not all necessary columns')
            raise IncorrectColumns

    def get_data(self) -> pd.DataFrame:
        """
        Get data from file
        :return: fixed file contents
        """
        return self.file_contents

    def delete_columns(self, columns: Tuple[str, ...]) -> None:
        """
        Delete columns from file
        :param columns: columns to delete
        """
        self.logger.debug('Delete columns')
        self.file_contents = self.file_contents.drop(columns=list(columns)).reset_index(drop=True)

    def replace_spaces_to_none(self):
        """
        Replace spaces to None in file contents
        """
        self.file_contents = self.file_contents.replace(r'^\s*$', np.nan, regex=True)

    def fix_row_number_sequence(self):
        """
        Fix row number sequence. Fix an error, when row number after nine set to one
        """
        for index in range(1, len(self.file_contents)):
            if index > 0:
                if self.file_contents.iloc[index - 1, 0][:-1] == self.file_contents.iloc[index, 0][:-1]:
                    self.file_contents.iloc[index, 0] = self.file_contents.iloc[index - 1, 0].replace('0', '1')

                    if self.file_contents.iloc[index - 1, 0].endswith('9') and not self.file_contents.iloc[index, 0].endswith('0'):