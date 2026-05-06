from __future__ import annotations

from pathlib import Path
import sys

from .config import AppConfig
from .excel_writer import get_processed_sources, save_outputs
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

    processed_sources = get_processed_sources(config.full_output_path)
    pending_files = filter_pending_files(hwp_files, processed_sources)
    if not pending_files:
        print("no new or changed YYYYMMDD.hwp files.")
        return 0
    print(f"processing {len(pending_files)} new or changed files.")

    try:
        records, failed_files = extract_records(pending_files, config, extractor)

        print(f"extracted {len(records)} records from {len(pending_files)} files.")
        if not records:
            print("no records extracted; output files were not created.", file=sys.stderr)
            return 4

        if dry_run:
            print_preview(records)
            return 0

        print("saving excel files...")
        successful_files = [
            hwp_file for hwp_file in pending_files if hwp_file.name not in failed_files
        ]
        full_output_path = save_outputs(
            records=records,
            fields=config.fields,
            template_path=config.result_template,
            output_path=config.full_output_path,
            updated_files=successful_files,
        )
        sorted_output_path = save_outputs(
            records=records,
            fields=config.fields,
            template_path=config.result_template,
            output_path=config.sorted_output_path,
            updated_files=successful_files,
            address_prefixes=config.address_prefixes,
            address_field=config.address_field,
        )
        print(f"saved full weekly workbook: {full_output_path}")
        print(f"saved sorted weekly workbook: {sorted_output_path}")
        if failed_files:
            print(f"skipped {len(failed_files)} files: {', '.join(sorted(failed_files))}")
        return 0
    finally:
        close = getattr(extractor, "close", None)
        if callable(close):
            close()


def discover_hwp_files(data_dir: Path) -> list[Path]:
    return sorted(data_dir.glob(DATE_NAME_PATTERN))


def source_signature(path: Path) -> str:
    stat = path.stat()
    return f"{stat.st_size}:{stat.st_mtime_ns}"


def filter_pending_files(
    hwp_files: list[Path],
    processed_sources: dict[str, str],
) -> list[Path]:
    return [
        hwp_file
        for hwp_file in hwp_files
        if processed_sources.get(hwp_file.name) != source_signature(hwp_file)
    ]


def extract_records(
    hwp_files: list[Path],
    config: AppConfig,
    extractor: object,
) -> tuple[list[ExtractedRecord], set[str]]:
    records: list[ExtractedRecord] = []
    failed_files: set[str] = set()
    for hwp_file in hwp_files:
        print(f"reading {hwp_file.name}...")
        try:
            tables = extractor.extract_tables(hwp_file, config.table_index)
        except HwpExtractionError as exc:
            print(f"skipping {hwp_file.name}: {exc}", file=sys.stderr)
            failed_files.add(hwp_file.name)
            continue

        file_record_count = 0
        for table in tables:
            table_records = table_to_records(
                table=table,
                fields=config.fields,
                source_file=hwp_file,
                orientation=config.input_orientation,
                newline_replacement=config.newline_replacement,
            )
            records.extend(table_records)
            file_record_count += len(table_records)
        print(f"finished {hwp_file.name}: {file_record_count} records.")
    return records, failed_files


def print_preview(records: list[ExtractedRecord], limit: int = 5) -> None:
    for record in records[:limit]:
        print(record.values)
