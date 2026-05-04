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

        self.assertEqual([record.values["index"] for record in records], ["1", "2"])
        self.assertEqual(records[0].values["name"], "A name")
        self.assertEqual(records[1].values["name"], "B name")
        self.assertTrue(records[0].contains_any(["경북"]))
        self.assertFalse(records[1].contains_any(["경북", "경남"]))

    def test_table_to_records_inserts_blank_phone_when_old_form_has_no_phone_row(self) -> None:
        table = [
            ["■ 137178"],
            ["광주 광산소방서 청사 증축"],
            ["광주광역시 소방안전본부"],
            ["비우종건(본:062-265-1216)"],
            ["광주 광산구 하남산단1번로 13"],
            ["지상1층 ~ 지상3층"],
            ["토목공사 진행중"],
            ["연면적:341㎡"],
        ]
        fields = [
            "index",
            "공사명",
            "건축주",
            "설계회사",
            "주소",
            "전화번호",
            "공사규모",
            "진행사항",
            "기타",
        ]

        records = table_to_records(table, fields, Path("20240624.hwp"))

        self.assertEqual(records[0].values["index"], "137178")
        self.assertEqual(records[0].values["전화번호"], "")
        self.assertEqual(records[0].values["공사규모"], "지상1층 ~ 지상3층")
        self.assertEqual(records[0].values["진행사항"], "토목공사 진행중")

    def test_table_to_records_keeps_phone_when_new_form_has_phone_row(self) -> None:
        table = [
            ["■ 147853"],
            ["구리 갈매동 515-1"],
            ["개인"],
            ["탄허건축"],
            ["서울시 한강대로 205"],
            ["02-790-1708"],
            ["지하2층 ~ 지상10층"],
            ["실시설계 진행중"],
            ["연면적:8,423㎡"],
        ]
        fields = [
            "index",
            "공사명",
            "건축주",
            "설계회사",
            "주소",
            "전화번호",
            "공사규모",
            "진행사항",
            "기타",
        ]

        records = table_to_records(table, fields, Path("20260504.hwp"))

        self.assertEqual(records[0].values["index"], "147853")
        self.assertEqual(records[0].values["전화번호"], "02-790-1708")
        self.assertEqual(records[0].values["공사규모"], "지하2층 ~ 지상10층")


if __name__ == "__main__":
    unittest.main()
