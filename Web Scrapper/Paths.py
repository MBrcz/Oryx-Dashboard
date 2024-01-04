# Libraries
from os.path import join
from datetime import datetime
from enum import Enum


_PATH_TO_SOURCE: str = "\\".join([f for f in __file__.split('\\')[:-2]])


class Paths(Enum):
    """
    This Enums stores all paths that are valuable in the project.
    """

    PATH_TO_FULL_TABLE = join(_PATH_TO_SOURCE, "Data Source", "factEquipment Losses.xlsx")
    PATH_TO_ORYX_DATA = join(_PATH_TO_SOURCE, "Data Source", "Oryx Data")
    PATH_TO_NEW_ORYX_DATA = join(_PATH_TO_SOURCE, "Data Source", "Oryx Data",
                                 f"factEquipment Losses {datetime.today().strftime('%Y-%m-%d')}.xlsx") # NOQA
    PATH_TO_DIM_DATES = join(_PATH_TO_SOURCE, "Data Source", f"dimDates.xlsx")


# Will hold them here - for future updates, if it will be ever made.
#     DATE = datetime.date.today()
#     Russia = "https://www.oryxspioenkop.com/2022/02/attack-on-europe-documenting-equipment.html"
#     Ukraine = "https://www.oryxspioenkop.com/2022/02/attack-on-europe-documenting-ukrainian.html"
#
#     # DATE = datetime.date(2022, 8, 31)
#     # Russia = "https://web.archive.org/web/20220831081921/https://www.oryxspioenkop.com/2022/02/attack-on-europe-documenting-equipment.html"
#     # Ukraine = "https://web.archive.org/web/20220831204235/https://www.oryxspioenkop.com/2022/02/attack-on-europe-documenting-ukrainian.html"
#
#     # DATE = datetime.date(2023, 1, 1)
#     # Russia = "https://web.archive.org/web/20230101145320/https://www.oryxspioenkop.com/2022/02/attack-on-europe-documenting-equipment.html"
#     # Ukraine = "https://web.archive.org/web/20230101163200/https://www.oryxspioenkop.com/2022/02/attack-on-europe-documenting-ukrainian.html"
#
#     # DATE = datetime.date(2023, 4, 30)
#     # Russia = "https://web.archive.org/web/20230430044641/https://www.oryxspioenkop.com/2022/02/attack-on-europe-documenting-equipment.html"
#     # Ukraine = "https://web.archive.org/web/20230430050410/https://www.oryxspioenkop.com/2022/02/attack-on-europe-documenting-ukrainian.html"
#
#     # DATE = datetime.date(2023, 8, 31)
#     # Russia = "https://web.archive.org/web/20230831184022/https://www.oryxspioenkop.com/2022/02/attack-on-europe-documenting-equipment.html"
#     # Ukraine = "https://web.archive.org/web/20230831143710/https://www.oryxspioenkop.com/2022/02/attack-on-europe-documenting-ukrainian.html"
