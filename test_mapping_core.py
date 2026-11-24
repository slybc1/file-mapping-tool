from __future__ import annotations

import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from mapping_core import MappingConfig, MappingEngine, MappingCancelled


class MappingEngineTest(unittest.TestCase):
    def test_exact_names_multiple_wo_and_all_sources(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source, target = root / 'source', root / 'target'
            for sub, content in [('a', 'first'), ('b', 'second')]:
                (source / sub).mkdir(parents=True)
                (source / sub / 'Report.pdf').write_text(content)
            (source / 'c').mkdir()
            (source / 'c' / 'report.PDF').write_text('different case')
            (target / 'WO1').mkdir(parents=True)
            (target / 'WO1' / 'REPORT.PDF').write_text('existing')
            book = Workbook()
            book.active.title = 'Data'
            for row in [('WO', 'File'), ('WO1', 'Report.pdf'), ('WO2', 'Report.pdf'),
                        ('WO1', 'Report.pdf'), ('WO3', 'report.PDF'), ('WO4', 'report.pdf')]:
                book.active.append(row)
            book.save(root / 'input.xlsx')
            with MappingEngine(MappingConfig(root / 'input.xlsx', 'Data', 1, 2,
                    source, target, root / 'result.xlsx', include_subfolders=True)) as engine:
                preview = engine.precheck()
                self.assertEqual((preview.matched_rows, preview.not_found_rows, preview.conflict_rows), (4, 1, 0))
                result = engine.execute()
                self.assertEqual((result.files_moved, result.rows_mapped), (3, 4))
                self.assertEqual((target / 'WO1' / 'REPORT.PDF').read_text(), 'existing')
                self.assertEqual({p.read_text() for p in (target / 'WO1').iterdir()}, {'existing', 'first', 'second'})
                self.assertEqual({p.read_text() for p in (target / 'WO2').iterdir()}, {'first', 'second'})
                self.assertEqual((target / 'WO3' / 'report.PDF').read_text(), 'different case')
                self.assertFalse((target / 'WO4').exists())
                self.assertFalse(list(source.rglob('*.pdf')))

    def test_failed_delivery_keeps_source_and_successful_wo(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / 'source'
            source.mkdir()
            (source / 'A.txt').write_text('original')
            book = Workbook()
            book.active.title = 'Data'
            for row in [('WO', 'File'), ('WO1', 'A.txt'), ('WO2', 'A.txt')]:
                book.active.append(row)
            book.save(root / 'input.xlsx')
            with MappingEngine(MappingConfig(root / 'input.xlsx', 'Data', 1, 2,
                    source, root / 'target', root / 'result.xlsx')) as engine:
                engine.precheck()
                deliver = engine._deliver
                def fail_one(src, folder, move):
                    if folder.name == 'WO2':
                        raise PermissionError('test failure')
                    return deliver(src, folder, move)
                with patch.object(engine, '_deliver', side_effect=fail_one):
                    result = engine.execute()
                self.assertEqual((result.rows_mapped, result.rows_failed), (1, 1))
                self.assertTrue((source / 'A.txt').exists())
                self.assertEqual((root / 'target/WO1/A.txt').read_text(), 'original')

    def test_cancel_during_shared_file_retains_source(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / 'source'
            source.mkdir()
            (source / 'A.txt').write_text('original')
            book = Workbook()
            book.active.title = 'Data'
            for row in [('WO', 'File'), ('WO1', 'A.txt'), ('WO2', 'A.txt')]:
                book.active.append(row)
            book.save(root / 'input.xlsx')
            with MappingEngine(MappingConfig(root / 'input.xlsx', 'Data', 1, 2,
                    source, root / 'target', root / 'result.xlsx')) as engine:
                engine.precheck()
                deliver = engine._deliver
                def cancel_second(src, folder, move):
                    if folder.name == 'WO2':
                        raise MappingCancelled()
                    return deliver(src, folder, move)
                with patch.object(engine, '_deliver', side_effect=cancel_second):
                    result = engine.execute()
                self.assertTrue(result.cancelled)
                self.assertEqual((result.rows_mapped, result.rows_not_processed), (1, 1))
                self.assertTrue((source / 'A.txt').exists())

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


class GroupedMappingTest(unittest.TestCase):
    def setUp(self) -> None:
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory()))
        self.source = self.root / "source"
        self.source.mkdir()
        self.target = self.root / "target"
        self.result_path = self.root / "result.xlsx"

    def mapping(self, rows: list[tuple]) -> MappingEngine:
        input_path = self.root / "input.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Data"
        for row in rows:
            sheet.append(row)
        workbook.save(input_path)
        workbook.close()
        config = MappingConfig(input_path, "Data", 1, 2, self.source, self.target, self.result_path)
        return MappingEngine(config)

    def row_results(self) -> list[tuple[str, str]]:
        workbook = load_workbook(self.result_path, read_only=True)
        try:
            rows = workbook["Data"].iter_rows(values_only=True)
            self.assertEqual(next(rows)[-2:], ("Mapping Status", "Failure Reason"))
            return [(status, reason or "") for *_, status, reason in rows]
        finally:
            workbook.close()

    def test_routes_by_year_type_and_ctx_id_and_merges_repeated_rows(self) -> None:
        filenames = ("Report.pdf", "Notes.txt")
        for name in filenames:
            (self.source / name).write_text(name, encoding="utf-8")
        wo_2025 = self.target / "2025" / "WO" / "77"
        wo_2025.mkdir(parents=True)
        (wo_2025 / "Report.pdf").write_text("existing", encoding="utf-8")

        with self.mapping([
            ("WO_NO", "File", " file_create_datetime ", " gc_ctx_type_cd ", " ctx_id "),
            ("wrong-id-1", "Report.pdf", datetime(2025, 1, 1), "WO", 77),
            ("wrong-id-2", "Report.pdf", "2025-12-31 23:59:59.999", "WO", 77),
            ("wrong-id-3", "Notes.txt", "2025-12-31 23:59:59", "WO", 77),
            ("wrong-id-4", "Report.pdf", "2025-12-31 23:59:59", "PM", 77),
            ("wrong-id-5", "Report.pdf", datetime(2026, 1, 1), "WO", 77),
            ("wrong-id-6", "report.PDF", "2025-01-01 00:00:00", "WO", 77),
        ]) as engine:
            preview = engine.precheck()
            self.assertEqual(preview.folders_to_create, 3)
            self.assertEqual((preview.matched_rows, preview.not_found_rows), (5, 1))
            self.assertTrue(all((self.source / name).is_file() for name in filenames))
            result = engine.execute()

        self.assertEqual((result.files_moved, result.rows_mapped), (2, 5))
        self.assertEqual(self.row_results(), [("Mapped", "")] * 5 + [("Not found", "")])
        expected = {
            "2025/WO/77/Report.pdf": "existing",
            "2025/WO/77/Report_2.pdf": "Report.pdf",
            "2025/WO/77/Notes.txt": "Notes.txt",
            "2025/PM/77/Report.pdf": "Report.pdf",
            "2026/WO/77/Report.pdf": "Report.pdf",
        }
        actual = {
            path.relative_to(self.target).as_posix(): path.read_text(encoding="utf-8")
            for path in self.target.rglob("*") if path.is_file()
        }
        self.assertEqual(actual, expected)
        self.assertFalse(list(self.source.iterdir()))

    def test_delivery_status_is_independent_for_each_full_destination(self) -> None:
        (self.source / "Report.pdf").write_text("attachment", encoding="utf-8")
        blocked = self.target / "2025" / "PM" / "77"
        blocked.parent.mkdir(parents=True)
        blocked.write_text("existing file blocks directory", encoding="utf-8")
        with self.mapping([
            ("CTX_ID", "File", "FILE_CREATE_DATETIME", "GC_CTX_TYPE_CD"),
            ("77", "Report.pdf", "2025-03-01 12:00:00", "WO"),
            ("77", "Report.pdf", "2025-03-01 12:00:00", "PM"),
        ]) as engine:
            self.assertEqual(engine.precheck().folders_to_create, 2)
            result = engine.execute()

        self.assertEqual((result.files_moved, result.rows_mapped, result.rows_failed), (0, 1, 1))
        ok, failed = self.row_results()
        self.assertEqual(ok, ("Mapped", ""))
        self.assertEqual(failed[0], "Move failed")
        self.assertTrue(failed[1])
        self.assertTrue((self.target / "2025/WO/77/Report.pdf").is_file())
        self.assertTrue((self.source / "Report.pdf").is_file())

    def test_uses_selected_identifier_when_ctx_id_header_is_absent(self) -> None:
        names = ("good.pdf", "bad-date.pdf", "missing-type.pdf", "unsafe-type.pdf", "missing-id.pdf")
        for name in names:
            (self.source / name).touch()
        with self.mapping([
            ("WO_NO", "File", "FILE_CREATE_DATETIME", "GC_CTX_TYPE_CD"),
            ("42", "good.pdf", "2025-01-01 00:00:00", "WO"),
            ("42", "bad-date.pdf", "2025-02-29 00:00:00", "WO"),
            ("42", "missing-type.pdf", "2025-01-01 00:00:00", None),
            ("42", "unsafe-type.pdf", "2025-01-01 00:00:00", "../PM"),
            (None, "missing-id.pdf", "2025-01-01 00:00:00", "WO"),
        ]) as engine:
            preview = engine.precheck()
            self.assertEqual((preview.matched_rows, preview.invalid_rows), (1, 4))
            result = engine.execute()

        self.assertEqual((result.files_moved, result.rows_failed), (1, 4))
        self.assertTrue((self.target / "2025/WO/42/good.pdf").is_file())
        self.assertEqual({path.name for path in self.source.iterdir()}, set(names[1:]))
        first, *failed = self.row_results()
        self.assertEqual(first, ("Mapped", ""))
        for status, reason in failed:
            self.assertEqual(status, "Move failed")
            self.assertTrue(reason)

    def test_rejects_incomplete_or_ambiguous_grouping_headers(self) -> None:
        (self.source / "Report.pdf").touch()
        cases = (
            ("FILE_CREATE_DATETIME",),
            ("GC_CTX_TYPE_CD",),
            ("FILE_CREATE_DATETIME", "GC_CTX_TYPE_CD", " gc_ctx_type_cd "),
        )
        for headers in cases:
            with self.subTest(headers=headers), self.mapping([
                ("WO_NO", "File", *headers), ("77", "Report.pdf"),
            ]) as engine:
                with self.assertRaises(ValueError):
                    engine.precheck()
                self.assertTrue((self.source / "Report.pdf").is_file())
                self.assertFalse(self.target.exists())

    def test_windows_folder_conflicts_are_scoped_to_the_parent(self) -> None:
        names = ("type1.pdf", "type2.pdf", "id1.pdf", "id2.pdf", "other-year.pdf")
        for name in names:
            (self.source / name).touch()
        with self.mapping([
            ("CTX_ID", "File", "FILE_CREATE_DATETIME", "GC_CTX_TYPE_CD"),
            ("A", "type1.pdf", datetime(2025, 1, 1), "WO"),
            ("B", "type2.pdf", datetime(2025, 1, 1), "wo"),
            ("X", "id1.pdf", datetime(2025, 1, 1), "PM"),
            ("x", "id2.pdf", datetime(2025, 1, 1), "PM"),
            ("x", "other-year.pdf", datetime(2026, 1, 1), "PM"),
        ]) as engine:
            preview = engine.precheck()
            self.assertEqual((preview.conflict_rows, preview.matched_rows), (4, 1))
            engine.execute()

        self.assertTrue((self.target / "2026/PM/x/other-year.pdf").is_file())
        self.assertEqual({path.name for path in self.source.iterdir()}, set(names[:-1]))
        self.assertEqual([status for status, _ in self.row_results()], ["Mapping conflict"] * 4 + ["Mapped"])


if __name__ == "__main__":
    unittest.main()
