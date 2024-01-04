# Libraries
from datetime import datetime
from enum import Enum
import pandas as pd
from os.path import join
from os import listdir


"""
This module is responsible for holding && formatting of the fact Table to the Oryx Project. Also keeps constans
"""


class FactColumns(Enum):
    """
    This enum stores the most important constans related to the fact Table and it's existance EXCEPT PATHS.
    """

    # HEADERS START
    ENTRY_NUM = "Entry Num"
    TYPE = "Type"
    FULL_NAME = "Full Name"
    COUNTRY_FLAG = "Country Flag"
    DATE = "Date"
    PROOF_OF_LOSS = "Proof of loss"
    OLD_QUANTITY_FATE = "Quantity, Fate"
    QUANTITY = "Quantity"
    FATE = " Fate"  # This " " is IMPORTANT!
    SIDE_OF_CONFLICT = "Side of Conflict"
    # HEADERS END

    @classmethod
    def get_headers_before_formatting(cls) -> list[str]:
        """
        This class method gets all the headers of the FactTableEnums object.
        THOSE HEADERS ARE BEFORE THE GENERAL FORMATTING!

        :return: list[str]
        """

        return [
            cls.TYPE.value,
            cls.FULL_NAME.value,
            cls.COUNTRY_FLAG.value,
            cls.PROOF_OF_LOSS.value,
            cls.OLD_QUANTITY_FATE.value,
            cls.SIDE_OF_CONFLICT.value
        ]


class FactTable:
    def __init__(self, scrapped_data: list[str] | pd.DataFrame) -> None:
        """
        Constructor for fact table Object.

        :param scrapped_data: the data that will be transferred into the fact Table
        """
        if isinstance(scrapped_data, list):
            self.fact_table = pd.DataFrame(scrapped_data, columns=FactColumns.get_headers_before_formatting())
        elif isinstance(scrapped_data, pd.DataFrame):
            self.fact_table = scrapped_data

    def get_added_date(self) -> list[str]:
        """
        Finds and returns the date that has been added to the dataframe.

        :return: None
        """

        return list(self.fact_table[FactColumns.DATE.value].unique())[0]

    def format_fact_table(self) -> None:
        """
        This is a summary method that shall format the newly aquired factTable from Oryx Website

        :return: None (modifies self.fact_table attribute)
        """

        print("Formatting newly acquired fact table")
        self._split_headers()
        self._remove_wayback_machine_traces()
        self._add_new_columns()
        self._other_formats()

    def _split_headers(self) -> None:
        """
        Split the headers from Fate and Quantity column.

        :return: None
        """

        column_to_split = FactColumns.OLD_QUANTITY_FATE.value
        left_col = column_to_split.split(',')[0]
        right_col = column_to_split.split(",")[1]

        self.fact_table[[left_col, right_col]] = self.fact_table[column_to_split].str.split(',', expand=True)
        self.fact_table = self.fact_table.drop(column_to_split, axis=1)

    def _remove_wayback_machine_traces(self) -> None:
        """
        This method will remove any traces about the usage of wayback machine in the columns (urls)
        This is written by CHATGPT.

        :return: None, modifies fact_table
        """

        splitting_phrase = "https"

        # Create a boolean mask to identify cells containing the splitting phrase
        mask = self.fact_table.map(
            lambda cell_value: isinstance(cell_value, str) and splitting_phrase in cell_value)

        # Use the mask to split the relevant cells and update them
        self.fact_table[mask] = self.fact_table[mask].map(lambda cell_value: splitting_phrase +
                                                          cell_value.split(splitting_phrase)[-1]
                                                          if isinstance(cell_value, str) else cell_value)

    def _add_new_columns(self) -> None:
        """
        Adds new columns to the Dataframe: Date and Entry Num. Also reorders them.

        :return: None, changes attribute fact_table.
        """

        self.fact_table[FactColumns.ENTRY_NUM.value] = self.fact_table.reset_index().index + 1
        self.fact_table[FactColumns.DATE.value] = datetime.now().strftime("%Y-%m-%d")

        self.fact_table.insert(2, FactColumns.DATE.value, self.fact_table.pop(FactColumns.DATE.value))
        self.fact_table.insert(0, FactColumns.ENTRY_NUM.value, self.fact_table.pop(FactColumns.ENTRY_NUM.value))

    def _other_formats(self) -> None:
        """
        Makes other cosmetic formats to the Dataframe.

        :return: None (modifies fact_table attribute)
        """

        self.fact_table[FactColumns.FATE.value] = self.fact_table[FactColumns.FATE.value].apply(lambda x: x.strip())
        self.fact_table[FactColumns.QUANTITY.value] = self.fact_table[FactColumns.QUANTITY.value].apply(lambda x: int(x)) # NOQA

    def save_file(self, path: str, sheet_name: str = "Sheet1") -> None:
        """
        This method saves the file somewhere in the user's machine

        :param path: The place where the file will be saved as an excel file.
        :param sheet_name: the name of the sheet where the item will be stored.
        :return: None
        """

        self.fact_table.to_excel(path, sheet_name=sheet_name, index_label=False, index=False)

    def filter_date_by_list_of_dates(self, list_of_dates: list):
        """
        This method filters a whole dataframe by the list of dates.

        :return: None
        """

        mask = self.fact_table[FactColumns.DATE.value].isin(list_of_dates)
        self.fact_table = self.fact_table[mask]

    def reset_entry_index(self):
        """
        This method restarts the Entry Num indexes in whole dataframe.

        :return: None (modifies the fact_table)
        """

        self.fact_table = self.fact_table.reset_index(drop=True)
        self.fact_table[FactColumns.ENTRY_NUM.value] = self.fact_table.index + 1

    @classmethod
    def combine_many_frames_into_one(cls, path: str) -> pd.DataFrame:
        """
        Loops for all files in the directory and appends its content to the big dataframe

        :param path: The directory from which the looping procedure will start.
        :return: pd.Dataframe - probably a BIG dataframe
        """

        merged_df: pd.DataFrame = pd.DataFrame()
        for count, filename in enumerate(listdir(path)):
            file_path = join(path, filename)
            df = pd.read_excel(file_path)
            merged_df = pd.concat([merged_df, df])
            print(f"Merged {count} files out of {len(listdir(path))}")
        return merged_df
