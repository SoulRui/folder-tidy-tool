# -*- coding: utf-8 -*-
import os
import tempfile
import unittest

import openpyxl

import folder_tidy_tool as tool


ROOT = os.path.dirname(os.path.abspath(__file__))


class FileNumberMatchingTests(unittest.TestCase):
    def setUp(self):
        self.mapping = {
            "2": ["一级", "2"],
            "4": ["一级", "4"],
            "65": ["一级", "65"],
        }

    def test_ascii_and_chinese_parentheses_use_base_number(self):
        for filename in (
            "4 (1).JPG", "4(2).JPG", "4（3）.JPG", "4 （4）.JPG",
            "4 (1) (2).JPG", "4.jpg", "4.JPG",
        ):
            with self.subTest(filename=filename):
                self.assertEqual(tool.match_file_key(filename, self.mapping), "4")

    def test_parenthetical_number_is_never_used_as_fallback(self):
        self.assertIsNone(tool.match_file_key("1 (2).JPG", self.mapping))
        self.assertIsNone(tool.match_file_key("12 (2).JPG", self.mapping))
        self.assertIsNone(tool.match_file_key("12.JPG", self.mapping))

    def test_non_numeric_names_keep_contains_matching(self):
        mapping = {"A04": ["一级", "A04"]}
        self.assertEqual(tool.match_file_key("照片A04 (2).JPG", mapping), "A04")

    def test_directory_matching_function_is_unchanged(self):
        self.assertEqual(tool.match_key("编号4", self.mapping), "4")
        self.assertEqual(tool.match_key("4", self.mapping), "4")

    def test_explicit_plot_number_ignores_years_and_partial_numbers(self):
        filename = "新昌县2026年林草执法线索146号图斑（2023年国土三调）示意图.jpg"
        mapping = dict.fromkeys(("23", "46", "146"), [])
        self.assertEqual(tool.match_file_key(filename, mapping), "146")

        mapping_with_suffix = dict.fromkeys(("23号", "46号", "146号"), [])
        self.assertEqual(tool.match_file_key(filename, mapping_with_suffix), "146号")

    def test_unlisted_explicit_number_never_falls_back_to_other_values(self):
        filename = "新昌县2026年林草执法线索146号图斑（2023年国土三调）示意图.jpg"
        mapping = dict.fromkeys(("23", "46", "2023", "2026"), [])
        self.assertIsNone(tool.match_file_key(filename, mapping))

    def test_reported_names_only_match_their_complete_plot_numbers(self):
        expected = {
            "146": None, "151": None, "166": None, "177": None,
            "20": None, "221": "221", "22": "22", "231": "231",
            "26": None, "271": None, "36": None,
        }
        mapping = dict.fromkeys(("23", "16", "17", "46", "221", "22", "231"), [])
        for number, expected_key in expected.items():
            filename = f"新昌县2026年林草执法线索{number}号图斑（2023年国土三调）示意图.jpg"
            with self.subTest(number=number):
                self.assertEqual(tool.match_file_key(filename, mapping), expected_key)

    def test_temporal_values_are_not_business_numbers(self):
        mapping = dict.fromkeys(("2023", "23", "5", "6", "11", "2"), [])
        for name in ("项目2023年5月6日", "项目2023年5月6号", "项目2023-05-06", "项目20231102"):
            with self.subTest(name=name):
                self.assertIsNone(tool.match_key(name, mapping))

    def test_plain_number_matches_classification_key_with_suffix(self):
        mapping = {"146号": []}
        self.assertEqual(tool.match_file_key("146.JPG", mapping), "146号")

    def test_classify_moves_files_using_number_before_parentheses(self):
        with tempfile.TemporaryDirectory(dir=ROOT) as temp_dir:
            excel_path = os.path.join(temp_dir, "分类表.xlsx")
            work_dir = os.path.join(temp_dir, "待处理")
            os.makedirs(work_dir)

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["编号", "一级文件夹名", "二级文件夹名"])
            for number in ("2", "4", "65"):
                ws.append([number, "一级", number])
            wb.save(excel_path)

            filenames = (
                "4 (1).JPG", "4 (2).JPG", "4（3）.JPG",
                "65 (4).JPG", "1 (2).JPG", "12 (2).JPG",
            )
            for filename in filenames:
                with open(os.path.join(work_dir, filename), "wb") as file:
                    file.write(b"test")

            tool.classify(excel_path, work_dir, mode="files", action="move", log=lambda message: None)

            for filename in ("4 (1).JPG", "4 (2).JPG", "4（3）.JPG"):
                self.assertTrue(os.path.isfile(os.path.join(work_dir, "一级", "4", filename)))
            self.assertTrue(os.path.isfile(os.path.join(work_dir, "一级", "65", "65 (4).JPG")))
            for filename in ("1 (2).JPG", "12 (2).JPG"):
                self.assertTrue(os.path.isfile(os.path.join(work_dir, tool.FAIL_DIR_NAME, filename)))
            self.assertFalse(os.path.isfile(os.path.join(work_dir, "一级", "2", "1 (2).JPG")))
            self.assertFalse(os.path.isfile(os.path.join(work_dir, "一级", "2", "12 (2).JPG")))

    def test_all_current_photo_names(self):
        excel_path = os.path.join(ROOT, "局里文件夹分类方式.xlsx")
        work_dir = os.path.join(ROOT, "待处理的文件")
        mapping = tool.load_classification(excel_path, log=lambda message: None)
        filenames = sorted(
            name for name in os.listdir(work_dir)
            if os.path.isfile(os.path.join(work_dir, name))
        )
        if not filenames:
            self.skipTest("未提供实际照片文件，跳过本地数据核验")
        self.assertEqual(len(filenames), 324)
        matched = 0
        unmatched = 0
        for filename in filenames:
            base = tool.strip_trailing_parenthetical(os.path.splitext(filename)[0])
            expected = base if base in mapping else None
            actual = tool.match_file_key(filename, mapping)
            with self.subTest(filename=filename):
                self.assertEqual(actual, expected)
            if actual is None:
                unmatched += 1
            else:
                matched += 1
        print(f"实际照片文件名核验：共 {len(filenames)} 个，匹配 {matched} 个，按表未匹配 {unmatched} 个")


if __name__ == "__main__":
    unittest.main(verbosity=2)
