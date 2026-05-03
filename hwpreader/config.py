from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path


@dataclass(frozen=True)
class AppConfig:
    table_index: int | None
    input_orientation: str
    newline_replacement: str
    address_prefixes: list[str]
    address_field: str
    fields: list[str]
    result_template: Path
    full_output_dir: Path
    sorted_output_dir: Path


def load_config(path: Path) -> AppConfig:
    raw = json.loads(path.read_text(encoding="utf-8"))
    base_dir = path.parent.parent
    return AppConfig(
        table_index=_load_table_index(raw.get("table_index", 0)),
        input_orientation=str(raw.get("input_orientation", "columns")),
        newline_replacement=str(raw.get("newline_replacement", " ")),
        address_prefixes=[
            str(item)
            for item in raw.get("address_prefixes", raw.get("keywords", []))
        ],
        address_field=str(raw.get("address_field", "주소")),
        fields=[str(item) for item in raw["fields"]],
        result_template=_resolve(base_dir, raw["result_template"]),
        full_output_dir=_resolve(base_dir, raw["full_output_dir"]),
        sorted_output_dir=_resolve(base_dir, raw["sorted_output_dir"]),
    )


def _resolve(base_dir: Path, value: str) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    return base_dir / path


def _load_table_index(value: object) -> int | None:
    if isinstance(value, str) and value.lower() == "all":
        return None
    return int(value)
