# Libraries
from bs4 import BeautifulSoup
from requests import get

# User defined Scripts
from Entry import Entry


"""
Scrapper object, it's main goal is to get the data from the websites and store it as a list of Entry objects.
"""


class Scrapper:
    COUNTRY: str | None = None
    DO_FORMAT = True  # Answers if user wants to get the raw data from web, or formatted one.

    def __init__(self):
        """
        Constructor for Scrapper object
        """
        self._url: str | None = None
        self._soup: BeautifulSoup | None = None
        self._result: list[Entry] = []

    @property
    def get_result(self) -> list[Entry]:
        return self._result

    def scrape_web_content(self) -> None:
        """
        This method searches for the equipment category in the website and then all bound to them equipment items.

        :return:
        """

        def create_entry() -> None:
            """
            Adds an entry to current loop iteration.
            :return: None
            """

            entry = Entry.create_instance(category_name, equipment_name, equipment_parent_country_flag,
                                          proof_url, entry_value, self.COUNTRY)

            if self.DO_FORMAT:
                entry.format_entry()
                entry.format_subcategory_name()

            self._result.append(entry)

        category_tag_num: int = 3  # See more upward.
        tag_decrement_value: int = 2

        while True:
            try:
                category_tag: str = self._soup.select("h3")[category_tag_num] # NOQA - type hint err
                category_name: str = category_tag.text.split('(')[0].strip() # NOQA - type hint err

                if category_name == "":
                    category_tag_num += 1  # This is the medicine for random 2 empty <h3> tags in the middle of
                                           # Ukrainian losses website.
                    tag_decrement_value += 1
                    continue  # Skip the loop

                subcategory_tag = self._soup.select("ul")[category_tag_num - tag_decrement_value]

                # This is stupid fix for wayback machine content grabbing problem
                if "<ul id=\"social\">" in str(subcategory_tag):
                    tag_decrement_value -= 1
                    subcategory_tag = self._soup.select("ul")[category_tag_num - tag_decrement_value]

                subcategory_tag = subcategory_tag.find_all("li")

                for count, subcategory in enumerate(subcategory_tag):
                    equipment_name: str = subcategory.text.split(":")[0].strip().replace(u'\xa0', u' ')

                    equipment_parent_country_flag: str = subcategory.find_all("img")[0]['src']

                    entry_value: str = ""
                    for oryx_entry in subcategory.find_all('a'):

                        proof_url: str = oryx_entry['href']
                        entry_value = entry_value + oryx_entry.text

                        # If there no brackets at the verges of string, add next string to the entry value.
                        # Prevents from code anomalies like this one:
                        # <a href="https://postimg.cc/q6CYJkkd">(2, </a>
                        # <a href="https://postimg.cc/q6CYJkkd">damaged</a>
                        # <a href="https://postimg.cc/q6CYJkkd">)</a>
                        # This should be counted as one entry only.

                        if entry_value[0] == "(" and entry_value[-1] == ")" and self.DO_FORMAT:
                            create_entry()

                        elif not self.DO_FORMAT:
                            create_entry()

                        entry_value = ""

                print(f"Added category: {category_name}, completed {category_tag_num}")
                category_tag_num += 1

            except IndexError as e:
                print("-----------------------------------")
                print(f"Exiting the loop due to: {e}")
                print(f"Category Tag Sum: {category_tag_num}")
                print("-----------------------------------")
                print('\n')
                break

    def open_website_content(self) -> None:
        """
        Sets the local soup attribute to the content of target website.

        :return: None
        """

        query = get(self._url)
        print(f"Opening website: {self._url}, response is: {query}")
        self._soup = BeautifulSoup(query.content, "html.parser")

    @classmethod
    def create_instance(cls, url: str):
        """Factory method that allows to create the instance of the class"""

        instance: Scrapper = Scrapper()
        instance._url = url
        return instance
