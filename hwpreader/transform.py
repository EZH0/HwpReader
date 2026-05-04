from __future__ import annotations

import re
from pathlib import Path

from .models import ExtractedRecord, Table


_SPACE_RE = re.compile(r"[ \t\r\f\v]+")
_INDEX_RE = re.compile(r"\d+")


def table_to_records(
    table: Table,
    fields: list[str],
    source_file: Path,
    orientation: str = "columns",
    newline_replacement: str = " ",
) -> list[ExtractedRecord]:
    if orientation != "columns":
        raise ValueError(f"Unsupported input_orientation: {orientation}")

    normalized = _normalize_table(table, newline_replacement)
    if not normalized:
        return []

    max_cols = max(len(row) for row in normalized)
    records: list[ExtractedRecord] = []

    for col_idx in range(max_cols):
        column = [_cell_at(row, col_idx) for row in normalized]
        if not any(column):
            continue
        column = _align_missing_optional_fields(column, fields)

        values: dict[str, str] = {}
        for field_idx, field_name in enumerate(fields):
            value = column[field_idx] if field_idx < len(column) else ""
            values[field_name] = _clean_field_value(field_name, value)

        if _looks_like_record(values):
            records.append(ExtractedRecord(source_file=source_file, values=values))

    return records


def _normalize_table(table: Table, newline_replacement: str) -> Table:
    return [
        [_clean_cell(cell, newline_replacement) for cell in row]
        for row in table
        if any(_clean_cell(cell, newline_replacement) for cell in row)
    ]


def _clean_cell(value: object, newline_replacement: str) -> str:
    if value is None:
        return ""
    text = str(value).replace("\n", newline_replacement)
    text = _SPACE_RE.sub(" ", text)
    return text.strip()


def _cell_at(row: list[str], index: int) -> str:
    return row[index] if index < len(row) else ""


def _looks_like_record(values: dict[str, str]) -> bool:
    return any(value for value in values.values())


def _align_missing_optional_fields(column: list[str], fields: list[str]) -> list[str]:
    if "전화번호" not in fields or "공사규모" not in fields:
        return column

    phone_index = fields.index("전화번호")
    if len(column) <= phone_index:
        return column

    if _looks_like_phone_number(column[phone_index]):
        return column

    if _looks_like_construction_scale(column[phone_index]):
        return column[:phone_index] + [""] + column[phone_index:]

    return column


def _clean_field_value(field_name: str, value: str) -> str:
    if field_name != "index":
        return value

    match = _INDEX_RE.search(value.replace(",", ""))
    return match.group(0) if match else value.replace("■", "").strip()


def _looks_like_phone_number(value: str) -> bool:
    return bool(re.search(r"\d{2,4}\s*[-)]\s*\d{3,4}", value))


def _looks_like_construction_scale(value: str) -> bool:
    compact = value.replace(" ", "")
    return any(keyword in compact for keyword in ("지하", "지상", "층", "㎡", "m²", "m2", "연면적"))
