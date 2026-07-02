from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from PyQt6.QtCore import QObject, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QIcon, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QSpinBox,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

import app


ROOT_DIR = Path(__file__).resolve().parent
LOGO_PATH = ROOT_DIR / "assets" / "passion_investment_logo.png"
DEFAULT_OUTPUT = ROOT_DIR / "outputs" / "lstk_vietnamnet.xlsx"


class ScrapeWorker(QObject):
    progress = pyqtSignal(str, object)
    finished = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, args: argparse.Namespace) -> None:
        super().__init__()
        self.args = args

    def run(self) -> None:
        try:
            result = app.run_scrape(self.args, progress_callback=self._emit_progress)
        except Exception as exc:  # noqa: BLE001 - surface worker errors in the UI.
            self.failed.emit(str(exc))
            return
        self.finished.emit(result)

    def _emit_progress(self, event: str, payload: dict[str, object]) -> None:
        self.progress.emit(event, payload)


class PassionInvestmentWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.thread: QThread | None = None
        self.worker: ScrapeWorker | None = None
        self.last_output_path = str(DEFAULT_OUTPUT)
        self._build_ui()

    def _build_ui(self) -> None:
        self.setWindowTitle("Passion Investment")
        self.setMinimumSize(980, 760)
        if LOGO_PATH.exists():
            self.setWindowIcon(QIcon(str(LOGO_PATH)))

        central = QWidget()
        self.setCentralWidget(central)

        outer = QVBoxLayout(central)
        outer.setContentsMargins(24, 24, 24, 24)
        outer.setSpacing(18)

        hero = QFrame()
        hero.setObjectName("hero")
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(28, 28, 28, 24)
        hero_layout.setSpacing(14)

        if LOGO_PATH.exists():
            logo_label = QLabel()
            pixmap = QPixmap(str(LOGO_PATH)).scaled(
                150,
                150,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
            logo_label.setPixmap(pixmap)
            logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            hero_layout.addWidget(logo_label, alignment=Qt.AlignmentFlag.AlignCenter)

        title = QLabel("PASSION INVESTMENT")
        title_font = QFont("Georgia", 24)
        title_font.setLetterSpacing(QFont.SpacingType.AbsoluteSpacing, 1.2)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setObjectName("titleLabel")

        subtitle = QLabel("Tool lay LSTK tu Vietnamnet vao Excel")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setObjectName("subtitleLabel")

        hero_layout.addWidget(title)
        hero_layout.addWidget(subtitle)
        outer.addWidget(hero)

        content_row = QHBoxLayout()
        content_row.setSpacing(18)
        outer.addLayout(content_row, stretch=1)

        controls_card = QFrame()
        controls_card.setObjectName("card")
        controls_layout = QVBoxLayout(controls_card)
        controls_layout.setContentsMargins(24, 24, 24, 24)
        controls_layout.setSpacing(18)

        controls_heading = QLabel("Scrape Setup")
        controls_heading.setObjectName("sectionTitle")
        controls_layout.addWidget(controls_heading)

        form = QGridLayout()
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(14)

        self.pages_spin = QSpinBox()
        self.pages_spin.setRange(1, 5000)
        self.pages_spin.setValue(500)

        self.stop_empty_spin = QSpinBox()
        self.stop_empty_spin.setRange(0, 100)
        self.stop_empty_spin.setValue(8)

        self.output_edit = QLineEdit(str(DEFAULT_OUTPUT))
        browse_button = QPushButton("Browse")
        browse_button.clicked.connect(self._browse_output)

        self.refresh_checkbox = QCheckBox("Refresh downloaded pages instead of cache")

        self.legacy_checkbox = QCheckBox("Search older free-text news before 31/5/2023")
        self.legacy_from_year_spin = QSpinBox()
        self.legacy_from_year_spin.setRange(2015, 2023)
        self.legacy_from_year_spin.setValue(2017)
        self.legacy_to_year_spin = QSpinBox()
        self.legacy_to_year_spin.setRange(2015, 2023)
        self.legacy_to_year_spin.setValue(2022)
        self.legacy_period_combo = QComboBox()
        self.legacy_period_combo.addItem("Monthly search", "month")
        self.legacy_period_combo.addItem("Daily search", "day")

        form.addWidget(self._field_label("Pages to scan"), 0, 0)
        form.addWidget(self.pages_spin, 0, 1)
        form.addWidget(self._field_label("Stop after empty pages"), 1, 0)
        form.addWidget(self.stop_empty_spin, 1, 1)
        form.addWidget(self._field_label("Output Excel file"), 2, 0)
        form.addWidget(self.output_edit, 2, 1)
        form.addWidget(browse_button, 2, 2)
        form.addWidget(self.refresh_checkbox, 3, 0, 1, 3)
        form.addWidget(self.legacy_checkbox, 4, 0, 1, 3)
        form.addWidget(self._field_label("Legacy from year"), 5, 0)
        form.addWidget(self.legacy_from_year_spin, 5, 1)
        form.addWidget(self._field_label("Legacy to year"), 6, 0)
        form.addWidget(self.legacy_to_year_spin, 6, 1)
        form.addWidget(self._field_label("Legacy search pace"), 7, 0)
        form.addWidget(self.legacy_period_combo, 7, 1)

        controls_layout.addLayout(form)

        button_row = QHBoxLayout()
        button_row.setSpacing(12)

        self.run_button = QPushButton("Start Scraping")
        self.run_button.setObjectName("primaryButton")
        self.run_button.clicked.connect(self._start_scrape)

        self.open_folder_button = QPushButton("Open Output Folder")
        self.open_folder_button.clicked.connect(self._open_output_folder)

        button_row.addWidget(self.run_button)
        button_row.addWidget(self.open_folder_button)
        controls_layout.addLayout(button_row)
        controls_layout.addStretch(1)

        status_card = QFrame()
        status_card.setObjectName("card")
        status_layout = QVBoxLayout(status_card)
        status_layout.setContentsMargins(24, 24, 24, 24)
        status_layout.setSpacing(14)

        status_heading = QLabel("Live Progress")
        status_heading.setObjectName("sectionTitle")
        status_layout.addWidget(status_heading)

        self.stage_label = QLabel("Ready")
        self.stage_label.setObjectName("statusPill")
        status_layout.addWidget(self.stage_label)

        self.page_progress = QProgressBar()
        self.page_progress.setFormat("Page scan: %v / %m")
        self.page_progress.setRange(0, 100)
        self.page_progress.setValue(0)

        self.article_progress = QProgressBar()
        self.article_progress.setFormat("Articles: %v / %m")
        self.article_progress.setRange(0, 100)
        self.article_progress.setValue(0)

        self.summary_label = QLabel("Waiting for launch")
        self.summary_label.setWordWrap(True)
        self.summary_label.setObjectName("summaryLabel")

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("Scrape updates will appear here...")

        status_layout.addWidget(self._field_label("Archive pages"))
        status_layout.addWidget(self.page_progress)
        status_layout.addWidget(self._field_label("Article scraping"))
        status_layout.addWidget(self.article_progress)
        status_layout.addWidget(self.summary_label)
        status_layout.addWidget(self.log_box, stretch=1)

        content_row.addWidget(controls_card, stretch=4)
        content_row.addWidget(status_card, stretch=6)

        self.setStyleSheet(
            """
            QWidget {
                background: #f4fbf6;
                color: #184d2f;
                font-family: "Segoe UI";
                font-size: 14px;
            }
            QMainWindow {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 1, y2: 1,
                    stop: 0 #eefaf0,
                    stop: 1 #dcefe1
                );
            }
            QFrame#hero, QFrame#card {
                background: rgba(255, 255, 255, 0.94);
                border: 2px solid #1d9b4a;
                border-radius: 24px;
            }
            QLabel#titleLabel {
                color: #13863f;
            }
            QLabel#subtitleLabel {
                color: #4c7c5f;
                font-size: 15px;
            }
            QLabel#sectionTitle {
                color: #11793a;
                font-size: 20px;
                font-weight: 600;
            }
            QLabel#statusPill {
                background: #e7f8eb;
                border: 1px solid #9ed3ae;
                border-radius: 14px;
                color: #11793a;
                padding: 8px 12px;
                font-weight: 600;
            }
            QLabel#summaryLabel {
                color: #2f6042;
                font-size: 13px;
            }
            QLabel.fieldLabel {
                color: #2f6042;
                font-size: 13px;
                font-weight: 600;
            }
            QLineEdit, QSpinBox, QTextEdit {
                background: #fcfffd;
                border: 1px solid #b7dcc3;
                border-radius: 12px;
                padding: 10px 12px;
            }
            QLineEdit:focus, QSpinBox:focus, QTextEdit:focus {
                border: 2px solid #18a14b;
            }
            QComboBox {
                background: #fcfffd;
                border: 1px solid #b7dcc3;
                border-radius: 12px;
                padding: 10px 12px;
            }
            QComboBox:focus {
                border: 2px solid #18a14b;
            }
            QCheckBox {
                spacing: 8px;
                color: #2f6042;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 1px solid #6db985;
                border-radius: 6px;
                background: white;
            }
            QCheckBox::indicator:checked {
                background: #18a14b;
            }
            QPushButton {
                background: #eff9f2;
                border: 1px solid #9ad3ab;
                border-radius: 14px;
                color: #16653a;
                font-weight: 600;
                padding: 12px 18px;
            }
            QPushButton:hover {
                background: #e0f5e7;
            }
            QPushButton#primaryButton {
                background: #179747;
                border: 1px solid #179747;
                color: white;
            }
            QPushButton#primaryButton:hover {
                background: #13863f;
            }
            QPushButton:disabled {
                background: #dfece3;
                color: #86a891;
                border-color: #d1e2d6;
            }
            QProgressBar {
                background: #edf6ef;
                border: 1px solid #b8dcc2;
                border-radius: 10px;
                text-align: center;
                min-height: 24px;
                color: #1c5d37;
            }
            QProgressBar::chunk {
                border-radius: 9px;
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #14a145,
                    stop: 1 #47bf71
                );
            }
            """
        )

    def _field_label(self, text: str) -> QLabel:
        label = QLabel(text)
        label.setProperty("class", "fieldLabel")
        label.setObjectName("fieldLabel")
        return label

    def _browse_output(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Choose Excel file",
            self.output_edit.text(),
            "Excel Workbook (*.xlsx)",
        )
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.output_edit.setText(path)

    def _start_scrape(self) -> None:
        if self.thread is not None:
            return

        output_path = self.output_edit.text().strip() or str(DEFAULT_OUTPUT)
        self.last_output_path = output_path
        self.log_box.clear()
        self.stage_label.setText("Running scrape")
        self.summary_label.setText("Preparing scraper...")
        self.page_progress.setValue(0)
        self.article_progress.setValue(0)
        self.page_progress.setMaximum(max(self.pages_spin.value(), 1))
        self.article_progress.setMaximum(1)
        self.run_button.setEnabled(False)

        args = argparse.Namespace(
            menu=False,
            pages=self.pages_spin.value(),
            start_page=0,
            base_url=app.BASE_TAG_URL,
            output=output_path,
            cache_dir=".cache/vietnamnet",
            refresh=self.refresh_checkbox.isChecked(),
            delay=0.15,
            timeout=25,
            max_articles=0,
            stop_after_empty=self.stop_empty_spin.value(),
            date_from=(
                f"{self.legacy_from_year_spin.value()}-01-01"
                if self.legacy_checkbox.isChecked()
                else ""
            ),
            date_to=(
                f"{self.legacy_to_year_spin.value()}-12-31"
                if self.legacy_checkbox.isChecked()
                else ""
            ),
            verbose=False,
            quiet=True,
            legacy=self.legacy_checkbox.isChecked(),
            legacy_from_year=self.legacy_from_year_spin.value(),
            legacy_to_year=self.legacy_to_year_spin.value(),
            legacy_period=self.legacy_period_combo.currentData(),
            legacy_max_results=10,
            legacy_max_articles=0,
            legacy_only=False,
        )

        self.thread = QThread()
        self.worker = ScrapeWorker(args)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._handle_progress)
        self.worker.finished.connect(self._handle_finished)
        self.worker.failed.connect(self._handle_failed)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self._cleanup_worker)
        self.thread.start()

    def _handle_progress(self, event: str, payload: object) -> None:
        data = payload if isinstance(payload, dict) else {}
        if event == "stage":
            self.stage_label.setText(str(data.get("message", "Working")))
            self.log_box.append(str(data.get("message", "")))
            return

        if event == "page_scan":
            done = int(data.get("pages_done", 0))
            total = max(int(data.get("pages_total", 1)), 1)
            self.page_progress.setMaximum(total)
            self.page_progress.setValue(min(done, total))
            message = (
                f"Page {int(data.get('page', 0))}: "
                f"{int(data.get('matches', 0))} matches, "
                f"{int(data.get('new_links', 0))} new links, "
                f"{int(data.get('total_links', 0))} total."
            )
            self.summary_label.setText(message)
            self.log_box.append(message)
            return

        if event == "links_ready":
            total_links = max(int(data.get("total_links", 0)), 1)
            self.article_progress.setMaximum(total_links)
            self.summary_label.setText(str(data.get("message", "")))
            self.log_box.append(str(data.get("message", "")))
            return

        if event == "legacy_period":
            done = int(data.get("done", 0))
            total = max(int(data.get("total", 1)), 1)
            self.page_progress.setMaximum(total)
            self.page_progress.setValue(min(done, total))
            message = (
                f"Legacy search {done}/{total}: {data.get('period', '')} "
                f"found {int(data.get('links', 0))} candidate links."
            )
            self.stage_label.setText("Searching legacy sources")
            self.summary_label.setText(message)
            self.log_box.append(message)
            return

        if event == "legacy_article":
            done = int(data.get("done", 0))
            total = max(int(data.get("total", done or 1)), 1)
            self.article_progress.setMaximum(total)
            self.article_progress.setValue(min(done, total))
            status = str(data.get("status", ""))
            rows = int(data.get("rows", 0))
            title = str(data.get("title", ""))
            message = f"Legacy article {done}/{total} | {status} | {rows} rows | {title}"
            self.summary_label.setText(message)
            self.log_box.append(message)
            return

        if event == "article":
            done = int(data.get("done", 0))
            total = max(int(data.get("total", 1)), 1)
            self.article_progress.setMaximum(total)
            self.article_progress.setValue(min(done, total))
            status = str(data.get("status", ""))
            rows = int(data.get("rows", 0))
            title = str(data.get("title", ""))
            message = f"{done}/{total} | {status} | {rows} rows | {title}"
            self.summary_label.setText(message)
            self.log_box.append(message)
            return

        if event == "finished":
            self.stage_label.setText("Completed")
            self.summary_label.setText(
                f"Finished with {int(data.get('rate_rows', 0))} exported rows."
            )
            self.log_box.append("Workbook export completed.")

    def _handle_finished(self, result: object) -> None:
        data = result if isinstance(result, dict) else {}
        self.last_output_path = str(data.get("output", self.last_output_path))
        self.run_button.setEnabled(True)
        self.stage_label.setText("Completed")
        self.summary_label.setText(
            f"Saved {int(data.get('rate_rows', 0))} rows from "
            f"{int(data.get('ok_articles', 0))} articles."
        )
        self.log_box.append(f"Saved file: {self.last_output_path}")
        QMessageBox.information(
            self,
            "Passion Investment",
            f"Scrape completed successfully.\n\nOutput:\n{self.last_output_path}",
        )

    def _handle_failed(self, message: str) -> None:
        self.run_button.setEnabled(True)
        self.stage_label.setText("Failed")
        self.summary_label.setText("The scrape stopped because of an error.")
        self.log_box.append(f"Error: {message}")
        QMessageBox.critical(self, "Passion Investment", message)

    def _cleanup_worker(self) -> None:
        if self.worker is not None:
            self.worker.deleteLater()
            self.worker = None
        if self.thread is not None:
            self.thread.deleteLater()
            self.thread = None
        self.run_button.setEnabled(True)

    def _open_output_folder(self) -> None:
        target = Path(self.output_edit.text().strip() or self.last_output_path)
        folder = target.parent if target.suffix else target
        folder.mkdir(parents=True, exist_ok=True)
        os.startfile(str(folder))


def main() -> int:
    qt_app = QApplication(sys.argv)
    window = PassionInvestmentWindow()
    window.show()
    return qt_app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
