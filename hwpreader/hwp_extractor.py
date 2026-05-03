from __future__ import annotations

from abc import ABC, abstractmethod
import contextlib
import io
from pathlib import Path
from typing import Iterator

from .models import Table


class HwpExtractionError(RuntimeError):
    pass


class HwpTableExtractor(ABC):
    name = "unknown"

    @abstractmethod
    def extract_table(self, path: Path, table_index: int) -> Table:
        raise NotImplementedError

    def extract_tables(self, path: Path, table_index: int | None) -> list[Table]:
        if table_index is not None:
            return [self.extract_table(path, table_index)]
        return [self.extract_table(path, 0)]


def create_extractor(backend: str = "auto") -> HwpTableExtractor:
    if backend not in {"auto", "pyhwpx", "com"}:
        raise ValueError(f"Unsupported backend: {backend}")

    if backend == "auto":
        return FallbackTableExtractor([PyhwpxTableExtractor, ComTableExtractor])

    if backend == "com":
        return ComTableExtractor()
    return PyhwpxTableExtractor()


class FallbackTableExtractor(HwpTableExtractor):
    name = "auto"

    def __init__(self, extractor_classes: list[type[HwpTableExtractor]]) -> None:
        self._extractor_classes = extractor_classes

    def extract_table(self, path: Path, table_index: int) -> Table:
        errors: list[str] = []
        for extractor_class in self._extractor_classes:
            try:
                extractor = extractor_class()
                return extractor.extract_table(path, table_index)
            except Exception as exc:
                errors.append(f"{extractor_class.name}: {exc}")
        detail = "; ".join(errors) if errors else "no backend tried"
        raise HwpExtractionError(f"Could not extract table from {path.name}. {detail}")

    def extract_tables(self, path: Path, table_index: int | None) -> list[Table]:
        errors: list[str] = []
        for extractor_class in self._extractor_classes:
            try:
                extractor = extractor_class()
                return extractor.extract_tables(path, table_index)
            except Exception as exc:
                errors.append(f"{extractor_class.name}: {exc}")
        detail = "; ".join(errors) if errors else "no backend tried"
        raise HwpExtractionError(f"Could not extract tables from {path.name}. {detail}")


class PyhwpxTableExtractor(HwpTableExtractor):
    name = "pyhwpx"

    def __init__(self) -> None:
        try:
            from pyhwpx import Hwp  # type: ignore
        except Exception as exc:
            raise HwpExtractionError("pyhwpx is not installed") from exc
        self._hwp_cls = Hwp

    def extract_table(self, path: Path, table_index: int) -> Table:
        hwp = None
        try:
            with _suppress_output():
                hwp = self._hwp_cls(visible=False)
                hwp.open(str(path))
                if hasattr(hwp, "get_into_nth_table") and hasattr(hwp, "table_to_df"):
                    hwp.get_into_nth_table(table_index)
                    return _ensure_table(_table_from_dataframe(hwp.table_to_df()), path)
                if hasattr(hwp, "get_table_text"):
                    text = hwp.get_table_text(table_index)
                    return _ensure_table(_parse_tabbed_table(str(text)), path)
            raise HwpExtractionError("pyhwpx table extraction method was not found")
        finally:
            if hwp is not None:
                _quit_hwp(hwp)

    def extract_tables(self, path: Path, table_index: int | None) -> list[Table]:
        if table_index is not None:
            return [self.extract_table(path, table_index)]

        hwp = None
        tables: list[Table] = []
        try:
            with _suppress_output():
                hwp = self._hwp_cls(visible=False)
                hwp.open(str(path))
                if not (hasattr(hwp, "get_into_nth_table") and hasattr(hwp, "table_to_df")):
                    raise HwpExtractionError("pyhwpx table extraction method was not found")

                index = 0
                while True:
                    selected = hwp.get_into_nth_table(index)
                    if not selected:
                        break
                    table = _ensure_table(_table_from_dataframe(hwp.table_to_df()), path)
                    tables.append(table)
                    index += 1
        finally:
            if hwp is not None:
                _quit_hwp(hwp)

        if not tables:
            raise HwpExtractionError(f"No tables were found in {path.name}")
        return tables


class ComTableExtractor(HwpTableExtractor):
    name = "com"

    def __init__(self) -> None:
        try:
            import win32com.client  # type: ignore
        except Exception as exc:
            raise HwpExtractionError("pywin32 is not installed") from exc
        self._win32 = win32com.client

    def extract_table(self, path: Path, table_index: int) -> Table:
        hwp = self._win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        try:
            hwp.Open(str(path))
            _move_to_table(hwp, table_index)
            return _ensure_table(_copy_selection_as_table(hwp), path)
        finally:
            _quit_hwp(hwp)

    def extract_tables(self, path: Path, table_index: int | None) -> list[Table]:
        if table_index is not None:
            return [self.extract_table(path, table_index)]

        tables: list[Table] = []
        index = 0
        while True:
            try:
                tables.append(self.extract_table(path, index))
            except HwpExtractionError:
                break
            index += 1
        if not tables:
            raise HwpExtractionError(f"No tables were found in {path.name}")
        return tables


def _move_to_table(hwp: object, table_index: int) -> None:
    # HWP stores tables as controls. This walks controls until the requested table is selected.
    for _ in range(table_index + 1):
        if not _run_action(hwp, "MoveSelNextCtrl"):
            raise HwpExtractionError(f"Table index {table_index} was not found")
    _run_action(hwp, "TableCellBlock")
    _run_action(hwp, "TableCellBlockExtend")
    _run_action(hwp, "TableCellBlockExtendAbs")


def _copy_selection_as_table(hwp: object) -> Table:
    _run_action(hwp, "Copy")
    text = _read_clipboard_text()
    table = _parse_tabbed_table(text)
    if not table:
        raise HwpExtractionError("Selected table was empty or clipboard copy failed")
    return table


def _read_clipboard_text() -> str:
    try:
        import tkinter as tk
    except Exception as exc:
        raise HwpExtractionError("tkinter clipboard access is unavailable") from exc

    root = tk.Tk()
    root.withdraw()
    try:
        return root.clipboard_get()
    finally:
        root.destroy()


def _parse_tabbed_table(text: str) -> Table:
    rows: Table = []
    for line in text.splitlines():
        cells = [cell.strip() for cell in line.split("\t")]
        if any(cells):
            rows.append(cells)
    return rows


def _table_from_dataframe(dataframe: object) -> Table:
    columns = [str(column).strip() for column in dataframe.columns]  # type: ignore[attr-defined]
    rows = dataframe.astype(str).values.tolist()  # type: ignore[attr-defined]
    if not rows or len(columns) < 2:
        return []

    table: Table = [[_normalize_cell(column) for column in columns[1:]]]
    for row in rows:
        table.append([_normalize_cell(cell) for cell in row[1:]])
    return table


def _normalize_cell(value: object) -> str:
    return str(value).replace("\r", " ").replace("\n", " ").strip()


def _ensure_table(table: Table, path: Path) -> Table:
    row_count = len(table)
    col_count = max((len(row) for row in table), default=0)
    if row_count < 2 or col_count < 2:
        raise HwpExtractionError(
            f"Extracted data does not look like a table from {path.name} "
            f"({row_count} rows x {col_count} columns)"
        )
    return table


@contextlib.contextmanager
def _suppress_output() -> Iterator[None]:
    stream = io.StringIO()
    with contextlib.redirect_stdout(stream), contextlib.redirect_stderr(stream):
        yield


def _run_action(hwp: object, action_name: str) -> bool:
    try:
        return bool(getattr(hwp, "Run")(action_name))
    except Exception:
        return False


def _quit_hwp(hwp: object) -> None:
    for method_name in ("FileClose", "Clear", "clear", "quit", "Quit", "FileQuit", "close", "Close"):
        method = getattr(hwp, method_name, None)
        if callable(method):
            try:
                with _suppress_output():
                    method()
            except Exception:
                pass
