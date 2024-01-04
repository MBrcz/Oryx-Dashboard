# User Defined Files.
from Scrapper import Scrapper
from FactTable import FactTable
from DimDates import DimDates
import Paths

"""
Starting file of whole operation.
"""

_WEBSITES = {
    "Russia": "https://www.oryxspioenkop.com/2022/02/attack-on-europe-documenting-equipment.html",
    "Ukraine": "https://www.oryxspioenkop.com/2022/02/attack-on-europe-documenting-ukrainian.html"
}


class Main:

    @staticmethod
    def scrape_website() -> list[str]:
        """
        Scraps the content of the website. See more: Configs and Scrapper.

        :return: list[str]
        """

        all_entries = []
        for owner, url in _WEBSITES.items():
            scrapper = Scrapper.create_instance(url)
            scrapper.COUNTRY = owner
            scrapper.open_website_content()
            scrapper.scrape_web_content()
            all_entries.extend(scrapper.get_result)

        return [entry.get_whole_entry() for entry in all_entries]

    @staticmethod
    def setup_the_new_entry_in_oryx_dir(list_of_entries: list[str]) -> list[str]:
        """
        Places the downloaded file in the directory where the oryx data is held.

        :return: list[str]
        """

        new_oryx_table = FactTable(list_of_entries)
        new_oryx_table.format_fact_table()
        new_oryx_table.save_file(Paths.Paths.PATH_TO_NEW_ORYX_DATA.value, sheet_name="Oryx")
        return new_oryx_table.get_added_date()

    @staticmethod
    def correct_the_dim_dates_file(date_to_add) -> list[str]:
        """
        Mades correction in dimDates file.
        WARNING!
        Do not stores any history-archive files in dates. Might cause issues later.

        :return: all unique dates from dimDates file.
        """

        dim_dates = DimDates(Paths.Paths.PATH_TO_DIM_DATES.value)
        dim_dates.cleanse_dates()
        dim_dates.append_new_date([date_to_add, _WEBSITES["Russia"], _WEBSITES["Ukraine"]])
        dim_dates.save_file()
        return dim_dates.get_all_unique_dates()

    @staticmethod
    def create_the_fact_equipment_losses(dates: list) -> None:
        """
        Handles the merge of the files and at the same time places the new factEquipment Losses file.

        :param dates: Dates that shall be filtered the new file.
        :return: None
        """

        print('---------------------------')
        print("Started merging the files")
        print('---------------------------')
        total_equipment_table = FactTable.combine_many_frames_into_one(Paths.Paths.PATH_TO_ORYX_DATA.value)

        fact_table = FactTable(total_equipment_table)
        fact_table.filter_date_by_list_of_dates(dates)
        fact_table.reset_entry_index()
        fact_table.save_file(Paths.Paths.PATH_TO_FULL_TABLE.value, "factEquipment Losses")
        print("Task completed.")

    @staticmethod
    def main() -> None:
        """
        Entry point of the scrapper. It:
        1) Gets all entries from Oryx website
        2) Sets up the new oryx data (for day)
        3) Corrects the dimDates table
        4) Creates the factEquipment Loss file.

        :return: None
        """

        entries = Main.scrape_website()
        print("Finished scrapping.")
        print(f"Creating {Paths.Paths.PATH_TO_NEW_ORYX_DATA.value} entry file.")
        date_added = Main.setup_the_new_entry_in_oryx_dir(entries)
        print(f"Correcting {Paths.Paths.PATH_TO_DIM_DATES.value} file.")
        all_dates = Main.correct_the_dim_dates_file(date_added)
        Main.create_the_fact_equipment_losses(all_dates)
        print("Special scrapping operation has been finished succesfully.\n"
              "Open Oryx Report.pbix and refresh it - it should work without any issues")


if __name__ == "__main__":
    Main.main()
