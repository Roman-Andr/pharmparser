from dataclasses import dataclass


@dataclass
class Settings:
    __slots__ = ["green", "red", "title", "data_title", "fileName", "colWidth", "cellWidth", "diffWidth"]
    green: str
    red: str
    title: str
    data_title: str
    fileName: str
    colWidth: int
    cellWidth: int
    diffWidth: int
