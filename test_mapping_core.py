from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from mapping_core import MappingConfig, MappingEngine


class MappingEngineTest(unittest.TestCase):
    def test_custom_columns_and_recursive_scan(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            workbook_path = root / "mapping.xlsx"
            source = root / "source"
            nested = source / "nested"
            target = root / "target"
            nested.mkdir(parents=True)
            target.mkdir()

            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Data"
            sheet.append(["Notes", "Exact File Name", "Unused", "Report Number"])
            sheet.append(["test", "Report001.txt", "x", "REPORT-31000001"])
            workbook.save(workbook_path)
            (nested / "Report001.txt").write_text("test", encoding="utf-8")

            config = MappingConfig(
                excel_path=workbook_path,
                sheet_name="Data",
                identifier_column=4,
                filename_column=2,
                source_folder=source,
                target_folder=target,
                result_path=root / "result.xlsx",
                include_subfolders=True,
            )
            engine = MappingEngine(config)
            try:
                preview = engine.precheck()
                self.assertEqual(preview.total_rows, 1)
                self.assertEqual(preview.source_files_scanned, 1)
                self.assertEqual(preview.matched_source_files, 1)

                summary = engine.execute()
                self.assertFalse(summary.cancelled)
                self.assertEqual(summary.files_moved, 1)
                self.assertTrue((target / "REPORT-31000001" / "Report001.txt").is_file())
                self.assertTrue(summary.result_path.is_file())
                self.assertTrue(summary.log_path.is_file())
            finally:
                engine.close()


if __name__ == "__main__":
    unittest.main()
