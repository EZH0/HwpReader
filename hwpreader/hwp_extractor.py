from __future__ import annotations

from abc import ABC, abstractmethod
import contextlib
import ctypes
import io
from pathlib import Path
import subprocess
import time
from typing import Iterator

from .models import Table


class HwpExtractionError(RuntimeError):
    pass


MAX_CONTROL_SCAN = 200


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
        return FallbackTableExtractor([ComTableExtractor, PyhwpxTableExtractor])

    if backend == "com":
        return ComTableExtractor()
    return PyhwpxTableExtractor()


class FallbackTableExtractor(HwpTableExtractor):
    name = "auto"

    def __init__(self, extractor_classes: list[type[HwpTableExtractor]]) -> None:
        self._extractor_classes = extractor_classes
        self._extractors: dict[type[HwpTableExtractor], HwpTableExtractor] = {}

    def extract_table(self, path: Path, table_index: int) -> Table:
        errors: list[str] = []
        for extractor_class in self._extractor_classes:
            try:
                extractor = self._get_extractor(extractor_class)
                return extractor.extract_table(path, table_index)
            except Exception as exc:
                errors.append(f"{extractor_class.name}: {exc}")
        detail = "; ".join(errors) if errors else "no backend tried"
        raise HwpExtractionError(f"Could not extract table from {path.name}. {detail}")

    def extract_tables(self, path: Path, table_index: int | None) -> list[Table]:
        errors: list[str] = []
        for extractor_class in self._extractor_classes:
            try:
                extractor = self._get_extractor(extractor_class)
                return extractor.extract_tables(path, table_index)
            except Exception as exc:
                errors.append(f"{extractor_class.name}: {exc}")
        detail = "; ".join(errors) if errors else "no backend tried"
        raise HwpExtractionError(f"Could not extract tables from {path.name}. {detail}")

    def close(self) -> None:
        for extractor in self._extractors.values():
            close = getattr(extractor, "close", None)
            if callable(close):
                close()

    def _get_extractor(
        self,
        extractor_class: type[HwpTableExtractor],
    ) -> HwpTableExtractor:
        if extractor_class not in self._extractors:
            self._extractors[extractor_class] = extractor_class()
        return self._extractors[extractor_class]


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
            import pythoncom  # type: ignore
            from win32com.client import dynamic  # type: ignore
        except Exception as exc:
            raise HwpExtractionError("pywin32 is not installed") from exc
        self._pythoncom = pythoncom
        self._dynamic = dynamic
        self._hwp: object | None = None
        self._owned_pids: set[int] = set()

    def extract_table(self, path: Path, table_index: int) -> Table:
        hwp = self._get_hwp()
        _register_file_path_check_module(hwp)
        hwp.Open(str(path))
        try:
            tables = _scan_tables(hwp, path, table_index)
        finally:
            _close_document(hwp)
        if len(tables) <= table_index:
            raise HwpExtractionError(f"Table index {table_index} was not found")
        return tables[table_index]

    def extract_tables(self, path: Path, table_index: int | None) -> list[Table]:
        if table_index is not None:
            return [self.extract_table(path, table_index)]

        hwp = self._get_hwp()
        _register_file_path_check_module(hwp)
        hwp.Open(str(path))
        try:
            tables = _scan_tables(hwp, path)
        finally:
            _close_document(hwp)

        if not tables:
            raise HwpExtractionError(f"No tables were found in {path.name}")
        return tables

    def close(self) -> None:
        if self._hwp is not None:
            _quit_hwp(self._hwp, self._owned_pids)
            self._hwp = None
            self._owned_pids = set()

    def _get_hwp(self) -> object:
        if self._hwp is None:
            before_pids = _hwp_process_ids()
            self._hwp = self._create_hwp_object()
            self._owned_pids = _hwp_process_ids() - before_pids
            window_pid = _hwp_window_process_id(self._hwp)
            if window_pid is not None and window_pid not in before_pids:
                self._owned_pids.add(window_pid)
            _register_file_path_check_module(self._hwp)
        return self._hwp

    def _create_hwp_object(self) -> object:
        clsid = self._pythoncom.CLSIDFromProgID("HWPFrame.HwpObject")
        dispatch = self._pythoncom.CoCreateInstance(
            clsid,
            None,
            self._pythoncom.CLSCTX_LOCAL_SERVER,
            self._pythoncom.IID_IDispatch,
        )
        return self._dynamic.Dispatch(dispatch)


def _move_to_table(hwp: object, table_index: int) -> None:
    # HWP stores tables as controls. This walks controls until the requested table is selected.
    for _ in range(table_index + 1):
        if not _run_action(hwp, "MoveSelNextCtrl"):
            raise HwpExtractionError(f"Table index {table_index} was not found")
    _run_action(hwp, "TableCellBlock")
    _run_action(hwp, "TableCellBlockExtend")
    _run_action(hwp, "TableCellBlockExtendAbs")


def _register_file_path_check_module(hwp: object) -> None:
    for module_name in (
        "FilePathCheckerModule",
        "FilePathCheckerModuleExample",
        "SecurityModule",
    ):
        try:
            if getattr(hwp, "RegisterModule")("FilePathCheckDLL", module_name):
                return
        except Exception:
            pass


def _close_document(hwp: object) -> None:
    for method_name in ("FileClose", "Clear", "clear", "close", "Close"):
        method = getattr(hwp, method_name, None)
        if callable(method):
            try:
                with _suppress_output():
                    method()
                return
            except Exception:
                pass


def _scan_tables(
    hwp: object,
    path: Path,
    table_index: int | None = None,
) -> list[Table]:
    tables: list[Table] = []
    for _ in range(MAX_CONTROL_SCAN):
        if not _run_action(hwp, "MoveSelNextCtrl"):
            break

        _run_action(hwp, "TableCellBlock")
        _run_action(hwp, "TableCellBlockExtend")
        _run_action(hwp, "TableCellBlockExtendAbs")

        with contextlib.suppress(HwpExtractionError):
            table = _ensure_table(_copy_selection_as_table(hwp), path)
            tables.append(table)
            if table_index is not None and len(tables) > table_index:
                break
    return tables


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


def _quit_hwp(hwp: object, owned_pids: set[int] | None = None) -> None:
    _close_document(hwp)
    if not owned_pids:
        return

    for method_name in ("Quit", "quit", "FileQuit"):
        method = getattr(hwp, method_name, None)
        if callable(method):
            try:
                with _suppress_output():
                    method()
                break
            except Exception:
                pass

    _terminate_owned_hwp_processes(owned_pids)


def _hwp_process_ids() -> set[int]:
    try:
        result = subprocess.run(
            [
                "tasklist",
                "/FI",
                "IMAGENAME eq hwp.exe",
                "/FO",
                "CSV",
                "/NH",
            ],
            capture_output=True,
            text=True,
            check=False,
            timeout=5,
        )
    except Exception:
        return set()

    process_ids: set[int] = set()
    for line in result.stdout.splitlines():
        parts = [part.strip().strip('"') for part in line.split('","')]
        if len(parts) >= 2 and parts[0].lower() == "hwp.exe":
            with contextlib.suppress(ValueError):
                process_ids.add(int(parts[1]))
    return process_ids


def _hwp_window_process_id(hwp: object) -> int | None:
    try:
        windows = getattr(hwp, "XHwpWindows")
        window = windows.Item(0)
    except Exception:
        return None

    hwnd = None
    for attr_name in ("HWND", "Hwnd", "hwnd"):
        with contextlib.suppress(Exception):
            hwnd = int(getattr(window, attr_name))
            break
    if not hwnd:
        return None

    pid = ctypes.c_ulong()
    ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return int(pid.value) if pid.value else None


def _terminate_owned_hwp_processes(owned_pids: set[int]) -> None:
    if not owned_pids:
        return

    deadline = time.monotonic() + 5
    while time.monotonic() < deadline:
        if not (owned_pids & _hwp_process_ids()):
            return
        time.sleep(0.2)

    for pid in sorted(owned_pids & _hwp_process_ids()):
        with contextlib.suppress(Exception):
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/T", "/F"],
                capture_output=True,
                text=True,
                check=False,
                timeout=5,
            )
