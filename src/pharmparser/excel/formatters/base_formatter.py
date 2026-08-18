from abc import ABC, abstractmethod

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from ...domain import PriceTable
from ...utils import Settings

MIN_STYLED_COLUMNS = 26
"""Style at least A..Z so sheets keep their familiar width even when narrow.

Bug B10 was the reverse of this: the formatters iterated ``string.ascii_uppercase``
and therefore stopped styling at Z, silently dropping column widths and conditional
formatting past ~13 pharmacies.
"""


class BaseFormatter(ABC):
    __slots__ = ["settings", "table"]

    def __init__(self, settings: Settings, table: PriceTable):
        self.settings = settings
        self.table = table

    @abstractmethod
    def format(self, sheet: Worksheet) -> None: ...

    def _set_column_widths(self, ws: Worksheet, widths: dict[int, int], default: int, used: int) -> None:
        """Apply ``widths`` by 1-based column index, filling the rest with ``default``."""
        for column in range(1, max(used, MIN_STYLED_COLUMNS) + 1):
            ws.column_dimensions[get_column_letter(column)].width = widths.get(column, default)
