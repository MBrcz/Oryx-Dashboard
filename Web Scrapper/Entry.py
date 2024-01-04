# Libraries
import re


class Entry:
    """
    Data class that will represent an Entry that user wants to write.
    Also consists of few formatting attribute methods (they are here for my personal convienence).
    """
    def __init__(self):
        """
        Constructor for the Entry data class.

        """
        self._category_name: str | None = None  # This is for instance a Tank, Armored Vehicle Unit, Armoured Infantry Unit etc... # NOQA
        self._subcategory_name: str | None = None  # The specified name of the category, like T-72 or BMP-1 etc...
        self._subcategory_country_name: str | None = None  # URL for an images of flags, like Russian or Soviet
        self._source_of_proof: str | None = None  # As name suggests
        self._entry: str | None = None  # Entry content, like (1, destroyed) or (1, captured) etc.
        self._team: str | None = None  # Shows who was the owner of the equipment.

    @property
    def category_name(self):
        return self._category_name

    @category_name.setter
    def category_name(self, value):
        self._category_name = value

    @property
    def subcategory_name(self):
        return self._subcategory_name

    @subcategory_name.setter
    def subcategory_name(self, value):
        self._subcategory_name = value

    @property
    def subcategory_country_name(self):
        return self._subcategory_country_name

    @subcategory_country_name.setter
    def subcategory_country_name(self, value):
        self._subcategory_country_name = value

    @property
    def source_of_proof(self):
        return self._source_of_proof

    @source_of_proof.setter
    def source_of_proof(self, value):
        self._source_of_proof = value

    @property
    def entry(self):
        return self._entry

    @entry.setter
    def entry(self, value):
        self._entry = value

    @property
    def team(self):
        return self._team

    @team.setter
    def team(self, value):
        self._team = value

    def format_subcategory_name(self):
        """Gets rid of numbers from subcategory name."""

        parts = self._subcategory_name.split(' ')
        self._subcategory_name = ' '.join(parts[1:])

    def format_entry(self) -> None:
        """Performs the following actions in the entry attribute:
        a) replaces "and" with "" but matches the whole word, for instance it would not touch "abANDoned" word
        b) replaces "," with ""
        b1) replaces ["(", ")"] with ""
        c) changes value to every numeric value in a string to 1,
        d) sums all the ocurrences of 1 into one value.
        e) removes all ones and places sum from d)

        : return None
        """

        # a)
        self._entry = re.sub(r'\band\b',"",  self._entry)  # Made by ChatGPT
        self._entry = re.sub(r'\b-\b', "", self._entry) # Removes all - phrases.

        # b)
        self._entry = self._entry.replace(',', "")

        # b1)
        self._entry = self._entry.replace('(', "")
        self._entry = self._entry.replace(')', "")

        # c)
        result = re.sub(r'\d+', '1', self._entry)

        # d)
        # This is corrected version, before converting 1 into a legitemate increment,
        # checks whether before 1 there was any valid string.
        # This is protection before entries like: (1, B-237 'Rostov-na-Donu', damaged beyond economical repair)
        # In old version, it would count it as 2, becouse 1 and 237 (converted to 1) are numbers.
        # Now it is not possible, becouse "B" blocks thisto happen.

        count_of_ones: int = 0
        trigger = False
        for num, char in enumerate(result):
            if char == str(1):
                for letter in range(num):
                    if result[letter].isalpha():
                        trigger = True
                    
                if trigger:
                    continue
                count_of_ones += 1

        # e)
        result = result.replace('1', '')
        result = str(count_of_ones) + ", " + result
        self._entry = result.strip()

    def get_whole_entry(self) -> list[str]:
        """
        A method that returns all attributes of the class as a list.

        :return: List[str]
        """

        return [
            self._category_name,
            self._subcategory_name,
            self._subcategory_country_name,
            self._source_of_proof,
            self._entry,
            self._team
        ]

    @classmethod
    def create_instance(cls, category_name: str, subcategory_name: str,
                        subcategory_country_name: str, source_of_proof: str,
                        entry: str, team: str):
        """Factory method that creates an instance of the entry class"""

        instance = Entry()
        instance.category_name = category_name
        instance.subcategory_name = subcategory_name
        instance.subcategory_country_name = subcategory_country_name
        instance.source_of_proof = source_of_proof
        instance.entry = entry
        instance.team = team

        return instance
