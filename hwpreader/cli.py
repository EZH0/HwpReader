from __future__ import annotations

import argparse
from pathlib import Path

from .config import load_config
from .pipeline import run_pipeline


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    root = Path.cwd()
    config = load_config(root / args.config)
    return run_pipeline(
        root=root,
        config=config,
        backend=args.backend,
        dry_run=args.dry_run,
    )


def _parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract HWP table data to Excel.")
    parser.add_argument("--config", default="resource/config.json")
    parser.add_argument("--backend", choices=["auto", "pyhwpx", "com"], default="auto")
    parser.add_argument("--dry-run", action="store_true")
    return parser.parse_args(argv)
