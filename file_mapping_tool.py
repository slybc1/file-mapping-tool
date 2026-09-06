from __future__ import annotations

import os
import queue
import sys
import threading
from datetime import datetime
from pathlib import Path

if "--smoke-test" in sys.argv:
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from openpyxl.utils import get_column_letter
from PySide6.QtCore import QTimer
from PySide6.QtGui import QCloseEvent, QFont
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from mapping_core import (
    ExecutionSummary,
    MappingCancelled,
    MappingConfig,
    MappingEngine,
    PreviewSummary,
    TaskControl,
    workbook_metadata,
)


APP_TITLE = "文件匹配与分类工具"


class FileMappingWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(920, 720)
        self.setMinimumSize(840, 640)

        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.control = TaskControl()
        self.engine: MappingEngine | None = None
        self.preview: PreviewSummary | None = None
        self.metadata: dict[str, list[tuple[int, str]]] = {}
        self.worker: threading.Thread | None = None

        self.excel_edit = QLineEdit()
        self.sheet_box = QComboBox()
        self.identifier_box = QComboBox()
        self.filename_box = QComboBox()
        self.source_edit = QLineEdit()
        self.target_edit = QLineEdit()
        self.result_edit = QLineEdit()
        self.case_insensitive = QCheckBox("文件名忽略英文字母大小写（Windows 推荐）")
        self.case_insensitive.setChecked(True)
        self.include_subfolders = QCheckBox(
            "包含待匹配文件夹下的所有子文件夹（文件很多时会增加扫描时间）"
        )
        self.include_subfolders.setChecked(False)

        self._build_ui()
        self._connect_configuration_changes()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._drain_events)
        self.timer.start(100)

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(24, 22, 24, 22)
        root.setSpacing(12)

        title = QLabel(APP_TITLE)
        title.setFont(QFont("Segoe UI", 18, QFont.Weight.DemiBold))
        root.addWidget(title)
        subtitle = QLabel("自由选择 Excel 列和待匹配文件夹；先预检查，确认后再移动文件。")
        subtitle.setObjectName("subtitle")
        root.addWidget(subtitle)

        setup = QGroupBox("匹配设置（每次运行均可重新选择）")
        form = QGridLayout(setup)
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)
        form.setColumnStretch(1, 1)

        self._path_row(form, 0, "Excel 原始数据", self.excel_edit, self._choose_excel)
        form.addWidget(QLabel("工作表"), 1, 0)
        form.addWidget(self.sheet_box, 1, 1, 1, 2)
        form.addWidget(QLabel("创建文件夹依据列"), 2, 0)
        form.addWidget(self.identifier_box, 2, 1, 1, 2)
        form.addWidget(QLabel("文件名所在列"), 3, 0)
        form.addWidget(self.filename_box, 3, 1, 1, 2)
        self._path_row(form, 4, "待匹配文件夹", self.source_edit, self._choose_source)
        form.addWidget(self.include_subfolders, 5, 1, 1, 2)
        self._path_row(form, 6, "分类结果文件夹", self.target_edit, self._choose_target)
        self._path_row(form, 7, "结果 Excel", self.result_edit, self._choose_result)
        form.addWidget(self.case_insensitive, 8, 1, 1, 2)
        root.addWidget(setup)

        actions = QHBoxLayout()
        self.precheck_button = QPushButton("1. 开始预检查")
        self.precheck_button.setObjectName("primary")
        self.precheck_button.clicked.connect(self._start_precheck)
        actions.addWidget(self.precheck_button)
        self.move_button = QPushButton("2. 确认并移动")
        self.move_button.clicked.connect(self._start_move)
        self.move_button.setEnabled(False)
        actions.addWidget(self.move_button)
        self.pause_button = QPushButton("暂停")
        self.pause_button.clicked.connect(self._toggle_pause)
        self.pause_button.setEnabled(False)
        actions.addWidget(self.pause_button)
        self.cancel_button = QPushButton("取消")
        self.cancel_button.clicked.connect(self._cancel)
        self.cancel_button.setEnabled(False)
        actions.addWidget(self.cancel_button)
        actions.addStretch(1)
        root.addLayout(actions)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        root.addWidget(self.progress_bar)
        self.status_label = QLabel("请先选择 Excel、列、待匹配文件夹和分类结果文件夹。")
        self.status_label.setObjectName("subtitle")
        root.addWidget(self.status_label)

        activity = QLabel("预检查结果与运行记录")
        activity.setFont(QFont("Segoe UI", 11, QFont.Weight.DemiBold))
        root.addWidget(activity)
        self.output = QPlainTextEdit()
        self.output.setReadOnly(True)
        self.output.setMinimumHeight(210)
        root.addWidget(self.output, 1)

        self.setStyleSheet(
            """
            QMainWindow, QWidget { background: #FFFFFF; color: #111827; font: 10pt 'Segoe UI'; }
            QLabel#subtitle { color: #4B5563; }
            QGroupBox { border: 1px solid #D1D5DB; border-radius: 6px; margin-top: 10px; padding-top: 12px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; font-weight: 600; }
            QLineEdit, QComboBox, QPlainTextEdit { border: 1px solid #CBD5E1; border-radius: 4px; padding: 6px; background: #FFFFFF; }
            QPlainTextEdit { background: #F8FAFC; }
            QPushButton { border: 1px solid #94A3B8; border-radius: 4px; padding: 7px 14px; background: #F8FAFC; }
            QPushButton:hover { background: #E2E8F0; }
            QPushButton#primary { background: #1F4E78; color: white; border-color: #1F4E78; font-weight: 600; }
            QPushButton#primary:hover { background: #173B5D; }
            QPushButton:disabled { color: #9CA3AF; background: #F3F4F6; border-color: #D1D5DB; }
            QProgressBar { border: 1px solid #CBD5E1; border-radius: 4px; text-align: center; height: 18px; }
            QProgressBar::chunk { background: #2563EB; }
            """
        )

    def _path_row(self, layout: QGridLayout, row: int, label: str, edit: QLineEdit, slot) -> None:
        layout.addWidget(QLabel(label), row, 0)
        layout.addWidget(edit, row, 1)
        button = QPushButton("选择…")
        button.clicked.connect(slot)
        layout.addWidget(button, row, 2)

    def _connect_configuration_changes(self) -> None:
        for edit in (self.excel_edit, self.source_edit, self.target_edit, self.result_edit):
            edit.textChanged.connect(self._invalidate_preview)
        self.sheet_box.currentTextChanged.connect(self._sheet_changed)
        self.identifier_box.currentTextChanged.connect(self._invalidate_preview)
        self.filename_box.currentTextChanged.connect(self._invalidate_preview)
        self.case_insensitive.stateChanged.connect(self._invalidate_preview)
        self.include_subfolders.stateChanged.connect(self._invalidate_preview)

    def _choose_excel(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self, "选择作为匹配依据的 Excel", "", "Excel 工作簿 (*.xlsx)"
        )
        if not selected:
            return
        self.excel_edit.setText(selected)
        path = Path(selected)
        self.result_edit.setText(
            str(path.with_name(f"{path.stem}_mapping_result_{datetime.now():%Y%m%d_%H%M%S}.xlsx"))
        )
        try:
            self.metadata = workbook_metadata(path)
        except Exception as exc:
            QMessageBox.critical(self, APP_TITLE, f"无法读取该 Excel。\n\n{exc}")
            return
        self.sheet_box.blockSignals(True)
        self.sheet_box.clear()
        self.sheet_box.addItems(list(self.metadata))
        self.sheet_box.blockSignals(False)
        self._refresh_columns()

    def _sheet_changed(self) -> None:
        self._refresh_columns()
        self._invalidate_preview()

    def _refresh_columns(self) -> None:
        headers = self.metadata.get(self.sheet_box.currentText(), [])
        for box in (self.identifier_box, self.filename_box):
            box.blockSignals(True)
            box.clear()
            for index, name in headers:
                box.addItem(f"{get_column_letter(index)} 列 · {name}", index)
            box.blockSignals(False)
        if headers:
            folder_index = self._suggest_column(
                headers, ("work order", "report", "number", "no.", " id", "编号", "工单", "报告")
            )
            file_index = self._suggest_column(
                headers, ("file name", "filename", "document name", "文件名", "文件名称")
            )
            self.identifier_box.setCurrentIndex(folder_index)
            self.filename_box.setCurrentIndex(
                file_index if file_index != folder_index else (1 if len(headers) > 1 else 0)
            )
        self._invalidate_preview()

    @staticmethod
    def _suggest_column(headers: list[tuple[int, str]], keywords: tuple[str, ...]) -> int:
        for position, (_index, name) in enumerate(headers):
            normalized = f" {name.casefold()} "
            if any(keyword in normalized for keyword in keywords):
                return position
        return 0

    def _choose_source(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "选择需要搜索和匹配文件的文件夹")
        if selected:
            self.source_edit.setText(selected)

    def _choose_target(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "选择分类结果保存文件夹")
        if selected:
            self.target_edit.setText(selected)

    def _choose_result(self) -> None:
        selected, _ = QFileDialog.getSaveFileName(
            self, "结果 Excel 另存为", "", "Excel 工作簿 (*.xlsx)"
        )
        if selected:
            if not selected.casefold().endswith(".xlsx"):
                selected += ".xlsx"
            self.result_edit.setText(selected)

    def _build_config(self) -> MappingConfig:
        required = {
            "Excel 原始数据": self.excel_edit.text(),
            "工作表": self.sheet_box.currentText(),
            "创建文件夹依据列": self.identifier_box.currentText(),
            "文件名所在列": self.filename_box.currentText(),
            "待匹配文件夹": self.source_edit.text(),
            "分类结果文件夹": self.target_edit.text(),
            "结果 Excel": self.result_edit.text(),
        }
        missing = [label for label, value in required.items() if not value.strip()]
        if missing:
            raise ValueError("请选择：" + "、".join(missing))
        return MappingConfig(
            excel_path=Path(self.excel_edit.text()),
            sheet_name=self.sheet_box.currentText(),
            identifier_column=int(self.identifier_box.currentData()),
            filename_column=int(self.filename_box.currentData()),
            source_folder=Path(self.source_edit.text()),
            target_folder=Path(self.target_edit.text()),
            result_path=Path(self.result_edit.text()),
            case_sensitive=not self.case_insensitive.isChecked(),
            include_subfolders=self.include_subfolders.isChecked(),
        )

    def _start_precheck(self) -> None:
        try:
            config = self._build_config()
        except Exception as exc:
            QMessageBox.critical(self, APP_TITLE, str(exc))
            return
        if self.engine is not None:
            self.engine.close()
        self.control = TaskControl()
        self.engine = MappingEngine(config, self._progress_event, self.control)
        self.preview = None
        self._set_busy(True)
        scope = "当前文件夹及所有子文件夹" if config.include_subfolders else "当前文件夹（不含子文件夹）"
        self.output.setPlainText(
            "预检查已开始，此阶段不会移动任何文件。\n"
            f"建文件夹依据：{self.identifier_box.currentText()}\n"
            f"匹配文件名依据：{self.filename_box.currentText()}\n"
            f"扫描范围：{scope}\n"
        )
        self.worker = threading.Thread(target=self._precheck_worker, daemon=True)
        self.worker.start()

    def _precheck_worker(self) -> None:
        assert self.engine is not None
        try:
            self.events.put(("preview", self.engine.precheck()))
        except MappingCancelled:
            self.events.put(("cancelled", "预检查已取消，没有改动任何文件。"))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def _start_move(self) -> None:
        if self.engine is None or self.preview is None:
            return
        prompt = (
            f"确定将 {self.preview.matched_source_files:,} 个匹配文件移动到 "
            f"{self.preview.folders_to_create:,} 个分类文件夹吗？\n\n"
            "原始 Excel 不会被修改；文件移动成功后将不再保留在原位置。"
        )
        answer = QMessageBox.warning(
            self,
            APP_TITLE,
            prompt,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.control = TaskControl()
        self.engine.control = self.control
        self._set_busy(True)
        self.output.appendPlainText("\n正式移动已开始。\n")
        self.worker = threading.Thread(target=self._move_worker, daemon=True)
        self.worker.start()

    def _move_worker(self) -> None:
        assert self.engine is not None
        try:
            self.events.put(("execution", self.engine.execute()))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def _toggle_pause(self) -> None:
        paused = self.control.toggle_pause()
        self.pause_button.setText("继续" if paused else "暂停")
        self.status_label.setText("已暂停" if paused else "正在继续")

    def _cancel(self) -> None:
        if (
            QMessageBox.question(self, APP_TITLE, "是否在当前文件操作完成后安全停止？")
            == QMessageBox.StandardButton.Yes
        ):
            self.control.cancel()
            self.status_label.setText("正在当前文件操作结束后安全取消")

    def _progress_event(self, phase: str, current: int, total: int | None, text: str) -> None:
        self.events.put(("progress", (phase, current, total, text)))

    def _drain_events(self) -> None:
        try:
            while True:
                event, payload = self.events.get_nowait()
                if event == "progress":
                    _phase, current, total, message = payload
                    self.status_label.setText(str(message))
                    if total and total > 0:
                        self.progress_bar.setRange(0, 100)
                        self.progress_bar.setValue(min(100, int(current * 100 / total)))
                    else:
                        self.progress_bar.setRange(0, 0)
                elif event == "preview":
                    self.progress_bar.setRange(0, 100)
                    self.progress_bar.setValue(100)
                    self.preview = payload
                    self._show_preview(self.preview)
                    self._set_busy(False)
                    self.move_button.setEnabled(bool(self.preview.matched_source_files))
                elif event == "execution":
                    self.progress_bar.setRange(0, 100)
                    self.progress_bar.setValue(100)
                    summary = payload
                    self._show_execution(summary)
                    self._set_busy(False)
                    self.move_button.setEnabled(bool(summary.cancelled and summary.rows_not_processed))
                elif event == "cancelled":
                    self.progress_bar.setRange(0, 100)
                    self.status_label.setText(str(payload))
                    self.output.appendPlainText(f"\n{payload}\n")
                    self._set_busy(False)
                elif event == "error":
                    self.progress_bar.setRange(0, 100)
                    self.status_label.setText("操作失败")
                    self.output.appendPlainText(f"\n错误：{payload}\n")
                    self._set_busy(False)
                    QMessageBox.critical(self, APP_TITLE, str(payload))
        except queue.Empty:
            pass

    def _show_preview(self, summary: PreviewSummary) -> None:
        volume_note = (
            "同一磁盘（移动速度快）"
            if summary.same_volume
            else "不同磁盘（实际为复制后删除，耗时较长）"
        )
        self.output.appendPlainText(
            "\n预检查完成\n"
            f"Excel 数据行：{summary.total_rows:,}\n"
            f"不重复的待匹配文件名：{summary.distinct_requested_files:,}\n"
            f"已扫描文件：{summary.source_files_scanned:,}\n"
            f"可匹配数据行：{summary.matched_rows:,}\n"
            f"准备移动的实际文件：{summary.matched_source_files:,}\n"
            f"未找到数据行：{summary.not_found_rows:,}\n"
            f"映射冲突数据行：{summary.conflict_rows:,}\n"
            f"无效数据行：{summary.invalid_rows:,}\n"
            f"涉及创建的分类文件夹：{summary.folders_to_create:,}\n"
            f"磁盘位置：{volume_note}\n"
            f"预检查耗时：{summary.elapsed_seconds:.1f} 秒\n"
        )
        self.status_label.setText("预检查完成。请核对数量，然后确认移动。")

    def _show_execution(self, summary: ExecutionSummary) -> None:
        state = "已取消" if summary.cancelled else "已完成"
        self.output.appendPlainText(
            f"\n处理{state}\n"
            f"已移动实际文件：{summary.files_moved:,}\n"
            f"标记为 Mapped 的数据行：{summary.rows_mapped:,}\n"
            f"失败或冲突数据行：{summary.rows_failed:,}\n"
            f"尚未处理数据行：{summary.rows_not_processed:,}\n"
            f"结果 Excel：{summary.result_path}\n"
            f"移动日志：{summary.log_path}\n"
            f"移动与生成结果耗时：{summary.elapsed_seconds:.1f} 秒\n"
        )
        message = "处理已取消，部分结果已经保存。" if summary.cancelled else "匹配处理完成，结果 Excel 和日志已经保存。"
        self.status_label.setText(message)
        QMessageBox.information(self, APP_TITLE, message)

    def _set_busy(self, busy: bool) -> None:
        self.precheck_button.setEnabled(not busy)
        self.pause_button.setEnabled(busy)
        self.pause_button.setText("暂停")
        self.cancel_button.setEnabled(busy)
        if busy:
            self.move_button.setEnabled(False)

    def _invalidate_preview(self) -> None:
        self.preview = None
        self.move_button.setEnabled(False)

    def closeEvent(self, event: QCloseEvent) -> None:
        if self.worker and self.worker.is_alive():
            self.control.cancel()
        elif self.engine is not None:
            self.engine.close()
        event.accept()


def main() -> int:
    application = QApplication(sys.argv)
    window = FileMappingWindow()
    if "--smoke-test" in sys.argv:
        window.show()
        application.processEvents()
        window.close()
        print("GUI smoke test passed")
        return 0
    window.show()
    return application.exec()


if __name__ == "__main__":
    raise SystemExit(main())
