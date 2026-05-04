from pathlib import Path
import tempfile
import unittest

from hwpreader.pipeline import filter_pending_files, source_signature


class PipelineTest(unittest.TestCase):
    def test_filter_pending_files_skips_unchanged_sources(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            data_dir = Path(temp_dir)
            processed = data_dir / "20260504.hwp"
            pending = data_dir / "20260505.hwp"
            processed.write_text("old", encoding="utf-8")
            pending.write_text("new", encoding="utf-8")

            files = [processed, pending]
            processed_sources = {processed.name: source_signature(processed)}

            self.assertEqual(filter_pending_files(files, processed_sources), [pending])


if __name__ == "__main__":
    unittest.main()
