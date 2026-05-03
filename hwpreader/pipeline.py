from __future__ import annotations

from pathlib import Path
import sys

from .config import AppConfig
from .excel_writer import save_outputs
from .hwp_extractor import HwpExtractionError, create_extractor
from .models import ExtractedRecord
from .transform import table_to_records


DATE_NAME_PATTERN = "????????.hwp"


def run_pipeline(
    root: Path,
    config: AppConfig,
    backend: str = "auto",
    dry_run: bool = False,
) -> int:
    hwp_files = discover_hwp_files(root / "data")
    if not hwp_files:
        print("data folder has no YYYYMMDD.hwp files.")
        return 0

    try:
        extractor = create_extractor(backend)
    except HwpExtractionError as exc:
        print(str(exc), file=sys.stderr)
        return 2

    try:
        records = extract_records(hwp_files, config, extractor)
    except HwpExtractionError as exc:
        print(str(exc), file=sys.stderr)
        return 3

    print(f"extracted {len(records)} records from {len(hwp_files)} files.")
    if not records:
        print("no records extracted; output files were not created.", file=sys.stderr)
        return 4

    if dry_run:
        print_preview(records)
        return 0

    print("saving excel files...")
    full_path, sorted_path = save_outputs(
        records=records,
        fields=config.fields,
        address_prefixes=config.address_prefixes,
        address_field=config.address_field,
        template_path=config.result_template,
        full_output_dir=config.full_output_dir,
        sorted_output_dir=config.sorted_output_dir,
    )
    print(f"saved full data: {full_path}")
    print(f"saved sorted data: {sorted_path}")
    return 0


def discover_hwp_files(data_dir: Path) -> list[Path]:
    return sorted(data_dir.glob(DATE_NAME_PATTERN))


def extract_records(
    hwp_files: list[Path],
    config: AppConfig,
    extractor: object,
) -> list[ExtractedRecord]:
    records: list[ExtractedRecord] = []
    for hwp_file in hwp_files:
        print(f"reading {hwp_file.name}...")
        tables = extractor.extract_tables(hwp_file, config.table_index)
        for table in tables:
            records.extend(
                table_to_records(
                    table=table,
                    fields=config.fields,
                    source_file=hwp_file,
                    orientation=config.input_orientation,
                    newline_replacement=config.newline_replacement,
                )
            )
    return records


def print_preview(records: list[ExtractedRecord], limit: int = 5) -> None:
    for record in records[:limit]:
        print(record.values)
