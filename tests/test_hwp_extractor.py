from types import SimpleNamespace
from unittest.mock import Mock, patch
import unittest

from hwpreader import hwp_extractor


class HwpExtractorTest(unittest.TestCase):
    def test_hwp_process_ids_parses_tasklist_csv_output(self) -> None:
        result = SimpleNamespace(
            stdout='"hwp.exe","1234","Console","1","10,000 K"\n'
            '"notepad.exe","9999","Console","1","5,000 K"\n'
        )

        with patch.object(hwp_extractor.subprocess, "run", return_value=result):
            self.assertEqual(hwp_extractor._hwp_process_ids(), {1234})

    def test_quit_hwp_does_not_quit_when_no_owned_process_was_found(self) -> None:
        hwp = SimpleNamespace(FileClose=Mock(), Quit=Mock())

        hwp_extractor._quit_hwp(hwp, set())

        hwp.FileClose.assert_called_once()
        hwp.Quit.assert_not_called()


if __name__ == "__main__":
    unittest.main()
