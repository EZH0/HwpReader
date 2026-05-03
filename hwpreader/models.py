from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


Table = list[list[str]]


@dataclass(frozen=True)
class ExtractedRecord:
    source_file: Path
    values: dict[str, str]

    def contains_any(self, keywords: list[str]) -> bool:
        haystack = " ".join(self.values.values())
        return any(keyword in haystack for keyword in keywords)
