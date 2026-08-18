"""VBA vocabulary for the generated sort and filter macros."""

from enum import Enum


class FilterCriteria(Enum):
    """An Excel ``AutoFilter`` criterion, spelled the way VBA expects it."""

    GREATER_THAN_ZERO = ">0"
    LESS_THAN_ZERO = "<0"
    GREATER_THAN_OR_EQUAL_ZERO = ">=0"
    LESS_THAN_OR_EQUAL_ZERO = "<=0"
    EQUAL_ZERO = "=0"


class SortOrder(Enum):
    """An Excel ``xlSortOrder`` constant."""

    ASCENDING = "xlAscending"
    DESCENDING = "xlDescending"
