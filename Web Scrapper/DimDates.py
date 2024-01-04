# Libraries
import pandas as pd

"""
This module holds the data about DimDates file.
"""


# The dates which won't be removed from the Date Table.
# Placed it here so it would be visible.
_ALLOWED_DATES = {
    # MONTHS : # DAYS
    8: 31,
    1: 1,
    4: 30,
    12: 31
}


class DimDatesColumns:
    DATE = "Date"
    RUSSIAN_SOURCE = "Russian Source"
    UKRAINIAN_SOURCE = "Ukrainian Source"


class DimDates:
    def __init__(self, path_to_file: str) -> None:
        """
        Constructor for DimDates object (table).

        :param path_to_file: path where the file is stored.
        :return None
        """
        self.path_to_file = path_to_file
        self.date_table = pd.read_excel(self.path_to_file)

    def get_all_unique_dates(self) -> list[str]:
        """
        This method takes all the unique dates from the Date Table.

        :return: list[str]
        """
        dates_as_string: list[str] = []
        dates = self.date_table[DimDatesColumns.DATE].unique()

        for date in dates:
            if isinstance(date, pd.Timestamp):
                dates_as_string.append(date.strftime("%Y-%m-%d"))
            else:
                dates_as_string.append(date)
        return dates_as_string

    def cleanse_dates(self) -> None:
        """
        This method loops through all the dates and removes all that are not counted as _ALLOWED.

        :return: None
        """

        self.date_table[DimDatesColumns.DATE] = pd.to_datetime(self.date_table[DimDatesColumns.DATE], errors='coerce')

        for count, date in enumerate(self.date_table['Date']):
            month, day = date.month, date.day
            if (month, day) not in _ALLOWED_DATES.items():
                self.date_table = self.date_table.drop(self.date_table.index[count])

    def append_new_date(self, list_values: list[str]) -> None:
        """
        Adds the new row to the DimDates dataframe.
        :param list_values: the values that will be added to a dataframe.

        :return: None
        """

        self.date_table.loc[len(self.date_table)] = list_values
        self.date_table.reset_index(drop=True)

    def save_file(self):
        """
        Saves the file (self) in the place where it was.

        :return: None
        """

        self.date_table.to_excel(self.path_to_file, index_label=False, index=False, sheet_name="dimDates")
