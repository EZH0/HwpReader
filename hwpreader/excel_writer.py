from __future__ import annotations

from datetime import datetime
from pathlib import Path
import shutil

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from .models import ExtractedRecord


SUMMARY_SHEET = "정리본"
PASTE_SHEET = "복붙용"


def save_outputs(
    records: list[ExtractedRecord],
    fields: list[str],
    address_prefixes: list[str],
    address_field: str,
    template_path: Path,
    full_output_dir: Path,
    sorted_output_dir: Path,
) -> tuple[Path, Path]:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    full_output_dir.mkdir(parents=True, exist_ok=True)
    sorted_output_dir.mkdir(parents=True, exist_ok=True)

    full_path = full_output_dir / f"full_data_{timestamp}.xlsx"
    sorted_path = sorted_output_dir / f"sorted_data_{timestamp}.xlsx"

    _write_workbook(records, fields, template_path, full_path)
    sorted_records = [
        record
        for record in records
        if _address_starts_with(record, address_field, address_prefixes)
    ]
    _write_workbook(sorted_records, fields, template_path, sorted_path)
    return full_path, sorted_path


def _address_starts_with(
    record: ExtractedRecord,
    address_field: str,
    address_prefixes: list[str],
) -> bool:
    address = record.values.get(address_field, "").strip()
    return any(address.startswith(prefix) for prefix in address_prefixes)


def _write_workbook(
    records: list[ExtractedRecord],
    fields: list[str],
    template_path: Path,
    output_path: Path,
) -> None:
    workbook = _load_or_create_workbook(template_path)
    summary = _get_or_create_sheet(workbook, SUMMARY_SHEET)
    paste = _get_or_create_sheet(workbook, PASTE_SHEET)

    _rewrite_summary_sheet(summary, records, fields)
    _rewrite_paste_sheet(paste, records, fields)
    workbook.save(output_path)


def _load_or_create_workbook(template_path: Path) -> Workbook:
    if template_path.exists():
        temp_path = template_path.parent / f".tmp_{template_path.name}"
        shutil.copy2(template_path, temp_path)
        try:
            return load_workbook(temp_path)
        finally:
            temp_path.unlink(missing_ok=True)
    return Workbook()


def _get_or_create_sheet(workbook: Workbook, name: str) -> Worksheet:
    if name in workbook.sheetnames:
        return workbook[name]
    return workbook.create_sheet(name)


def _rewrite_summary_sheet(
    sheet: Worksheet,
    records: list[ExtractedRecord],
    fields: list[str],
) -> None:
    _clear_sheet(sheet)
    headers = fields + ["source_file"]
    sheet.append(headers)
    for record in records:
        sheet.append([record.values.get(field, "") for field in fields] + [record.source_file.name])
    _style_header(sheet)
    _set_widths(sheet, headers)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions


def _rewrite_paste_sheet(
    sheet: Worksheet,
    records: list[ExtractedRecord],
    fields: list[str],
) -> None:
    _clear_sheet(sheet)
    for col_idx, record in enumerate(records, start=1):
        for row_idx, field in enumerate(fields, start=1):
            sheet.cell(row=row_idx, column=col_idx, value=record.values.get(field, ""))
    _style_header_row_as_index(sheet)
    _set_paste_widths(sheet, len(records))


def _clear_sheet(sheet: Worksheet) -> None:
    if sheet.max_row:
        sheet.delete_rows(1, sheet.max_row)


def _style_header(sheet: Worksheet) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill


def _style_header_row_as_index(sheet: Worksheet) -> None:
    fill = PatternFill("solid", fgColor="D9EAF7")
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = fill


def _set_widths(sheet: Worksheet, headers: list[str]) -> None:
    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, min(sheet.max_row, 80) + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            max_len = max(max_len, len(str(value or "")))
        sheet.column_dimensions[sheet.cell(row=1, column=col_idx).column_letter].width = min(max(max_len + 2, 10), 42)


def _set_paste_widths(sheet: Worksheet, count: int) -> None:
    for col_idx in range(1, count + 1):
        sheet.column_dimensions[sheet.cell(row=1, column=col_idx).column_letter].width = 24
