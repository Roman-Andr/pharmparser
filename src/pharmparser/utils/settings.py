from dataclasses import dataclass


@dataclass
class Settings:
    __slots__ = ["cellWidth", "colWidth", "diffWidth", "fileName", "green", "red", "title"]
    green: str
    red: str
    title: str
    fileName: str
    colWidth: int
    cellWidth: int
    diffWidth: int
