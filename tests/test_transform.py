from pathlib import Path
import unittest

from hwpreader.transform import table_to_records


class TransformTest(unittest.TestCase):
    def test_table_to_records_columns_and_cleaning(self) -> None:
        table = [
            ["■ 1", "■ 2"],
            ["A\nname", "B\t name"],
            ["경북 주소", "서울 주소"],
        ]
        fields = ["index", "name", "address"]

        records = table_to_records(table, fields, Path("20260430.hwp"))

        self.assertEqual([record.values["index"] for record in records], ["■ 1", "■ 2"])
        self.assertEqual(records[0].values["name"], "A name")
        self.assertEqual(records[1].values["name"], "B name")
        self.assertTrue(records[0].contains_any(["경북"]))
        self.assertFalse(records[1].contains_any(["경북", "경남"]))


if __name__ == "__main__":
    unittest.main()
