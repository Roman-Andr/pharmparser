from enum import Enum


class FilterCriteria(Enum):
    GREATER_THAN_ZERO = ">0"
    LESS_THAN_ZERO = "<0"
    GREATER_THAN_OR_EQUAL_ZERO = ">=0"
    LESS_THAN_OR_EQUAL_ZERO = "<=0"
    EQUAL_ZERO = "=0"
