from __future__ import annotations

import math
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from .constants import APP_TITLE, DEFAULT_TARGET_QUARTERS, EXPORT_DIR
from .database import Database
from .quarters import normalize_quarters, quarter_sort_key
from .services import FinancialCollectorService


def _to_float(value: Any) -> float | None:
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _format_number(value: Any, digits: int = 3) -> str:
    number = _to_float(value)
    if number is None:
        return "-"
    return f"{number:,.{digits}f}"


def _format_money_to_hundred_million(value: Any) -> str:
    number = _to_float(value)
    if number is None:
        return "-"
    return f"{number / 100000000:,.2f} 億"


def _calculate_target_price(final_eps: float | None, pe_ratio: float | None) -> float | None:
    if final_eps is None or pe_ratio is None:
        return None
    return final_eps * pe_ratio

class DraggableEpsCanvas(FigureCanvas):
    def __init__(self) -> None:
        self.figure = Figure(figsize=(7, 3.5), tight_layout=True)
        self.ax = self.figure.add_subplot(111)
        super().__init__(self.figure)

        self.quarters: list[str] = []
        self.values: list[float] = []
        self.editable_mask: list[bool] = []
        self.drag_index: int | None = None

        self.on_point_changed: Callable[[str, float], None] | None = None
        self.on_drag_finished: Callable[[str, float], None] | None = None

        self.mpl_connect("button_press_event", self._on_press)
        self.mpl_connect("motion_notify_event", self._on_motion)
        self.mpl_connect("button_release_event", self._on_release)

        self._draw_empty()

    def set_series(self, quarters: list[str], values: list[float | None], editable_mask: list[bool]) -> None:
        self.quarters = list(quarters)
        self.editable_mask = list(editable_mask)
        self.values = []

        for value, editable in zip(values, editable_mask):
            current = _to_float(value)
            if current is None and editable:
                self.values.append(math.nan)
            elif current is None:
                self.values.append(math.nan)
            else:
                self.values.append(current)

        self.drag_index = None
        self._redraw()

    def _rolling_total_eps(self) -> list[float | None]:
        """每個季度往前取 4 季（含自身）加總，作為 TTM EPS。"""
        result: list[float | None] = []
        for i in range(len(self.values)):
            window = self.values[max(0, i - 3): i + 1]
            valid = [v for v in window if not math.isnan(v)]
            if len(valid) == 4:
                result.append(round(sum(valid), 3))
            else:
                result.append(None)
        return result

    def _draw_empty(self) -> None:
        self.ax.clear()
        self.ax.set_title("EPS Trend", fontsize=12, fontweight="bold", color="#2b3f5c")
        self.ax.text(0.5, 0.5, "No data available", ha="center", va="center",
                     transform=self.ax.transAxes, fontsize=11, color="#888")
        self.ax.set_xticks([])
        self.ax.grid(alpha=0.2)
        self.figure.patch.set_facecolor("#f7fbff")
        self.draw_idle()

    def _redraw(self) -> None:
        self.ax.clear()
        if not self.quarters:
            self._draw_empty()
            return

        x_values = list(range(len(self.quarters)))
        finite_pairs = [(x, v) for x, v in zip(x_values, self.values) if not math.isnan(v)]

        # --- Quarterly EPS line ---
        if finite_pairs:
            fx, fv = zip(*finite_pairs)
            self.ax.plot(fx, fv, color="#3a86ff", linewidth=2.0,
                         alpha=0.85, label="Quarterly EPS", zorder=2)

        # --- Scatter points ---
        for idx, value in enumerate(self.values):
            if math.isnan(value):
                continue
            if self.editable_mask[idx]:
                color = "#ef476f"
                marker = "o"
                size = 80
            else:
                color = "#3a86ff"
                marker = "o"
                size = 60
            self.ax.scatter(idx, value, s=size, c=color,
                            edgecolor="#ffffff", linewidth=1.2, zorder=4, marker=marker)

        # --- TTM Total EPS line ---
        ttm = self._rolling_total_eps()
        ttm_x = [i for i, v in enumerate(ttm) if v is not None]
        ttm_y = [v for v in ttm if v is not None]
        if ttm_x:
            self.ax.plot(ttm_x, ttm_y, color="#f4a261", linewidth=1.8,
                         linestyle="--", alpha=0.9, label="TTM EPS (4Q Total)", zorder=3)
            for x, y in zip(ttm_x, ttm_y):
                self.ax.scatter(x, y, s=36, c="#f4a261",
                                edgecolor="#ffffff", linewidth=1.0, zorder=5)

        # --- EPS value labels ---
        for idx, value in enumerate(self.values):
            if not math.isnan(value):
                self.ax.annotate(
                    f"{value:.2f}",
                    xy=(idx, value),
                    xytext=(0, 10),
                    textcoords="offset points",
                    ha="center",
                    fontsize=8,
                    color="#2b3f5c",
                )

        # --- Axis & style ---
        all_values = [v for v in self.values if not math.isnan(v)] + [v for v in ttm_y]
        if not all_values:
            all_values = [0.0]
        min_val = min(all_values)
        max_val = max(all_values)
        padding = max(1.5, (max_val - min_val) * 0.2)
        self.ax.set_ylim(min_val - padding, max_val + padding)
        self.ax.set_xlim(-0.5, len(self.quarters) - 0.5)
        self.ax.set_xticks(x_values)
        self.ax.set_xticklabels(self.quarters, rotation=30, ha="right", fontsize=9)
        self.ax.set_title("EPS Trend  |  Red = estimated (draggable)", fontsize=11,
                          fontweight="bold", color="#2b3f5c", pad=10)
        self.ax.set_ylabel("EPS (NTD)", fontsize=9, color="#44566c")
        self.ax.grid(axis="y", alpha=0.2, linestyle="--")
        self.ax.spines[["top", "right"]].set_visible(False)
        self.ax.legend(loc="upper left", fontsize=8, framealpha=0.6)
        self.figure.patch.set_facecolor("#f7fbff")
        self.ax.set_facecolor("#f7fbff")
        self.draw_idle()

    def _on_press(self, event: Any) -> None:
        if event.inaxes is not self.ax or event.xdata is None or event.ydata is None:
            return
        if not self.quarters:
            return

        idx = int(round(event.xdata))
        if idx < 0 or idx >= len(self.values):
            return
        if not self.editable_mask[idx]:
            return

        point_value = self.values[idx]
        if math.isnan(point_value):
            point_value = float(event.ydata)
            self.values[idx] = point_value

        close_x = abs(event.xdata - idx) <= 0.45
        close_y = abs(event.ydata - point_value) <= max(0.8, abs(point_value) * 0.25)
        if close_x and close_y:
            self.drag_index = idx

    def _on_motion(self, event: Any) -> None:
        if self.drag_index is None or event.inaxes is not self.ax or event.ydata is None:
            return

        value = max(-20.0, min(200.0, float(event.ydata)))
        value = round(value, 3)
        if self.values[self.drag_index] == value:
            return

        self.values[self.drag_index] = value
        self._redraw()

        if self.on_point_changed:
            self.on_point_changed(self.quarters[self.drag_index], value)

    def _on_release(self, _event: Any) -> None:
        if self.drag_index is None:
            return
        idx = self.drag_index
        self.drag_index = None
        if self.on_drag_finished:
            self.on_drag_finished(self.quarters[idx], float(self.values[idx]))


class MainWindow(QMainWindow):
    def __init__(self, db: Database, collector: FinancialCollectorService) -> None:
        super().__init__()
        self.db = db
        self.collector = collector
        self._loading_eps_table = False

        self.setWindowTitle(APP_TITLE)
        self.resize(1540, 920)

        self._build_ui()
        self._apply_styles()
        self._refresh_all_views()

    def _build_ui(self) -> None:
        root = QWidget()
        root_layout = QHBoxLayout(root)
        root_layout.setContentsMargins(18, 18, 18, 18)
        root_layout.setSpacing(14)

        left_card, left_layout = self._create_card("公司清單")
        left_card.setFixedWidth(240)
        self.company_list = QListWidget()
        for profile in self.collector.company_profiles.values():
            item = QListWidgetItem(f"{profile.company_id}  {profile.name}")
            item.setData(Qt.UserRole, profile.company_id)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            self.company_list.addItem(item)
        left_layout.addWidget(self.company_list)

        left_buttons = QHBoxLayout()
        self.select_all_button = QPushButton("全選")
        self.clear_all_button = QPushButton("清除")
        left_buttons.addWidget(self.select_all_button)
        left_buttons.addWidget(self.clear_all_button)
        left_layout.addLayout(left_buttons)

        center_card, center_layout = self._create_card("操作台")
        self.tabs = QTabWidget()
        center_layout.addWidget(self.tabs)
        self._build_collect_tab()
        self._build_eps_tab()

        right_card, right_layout = self._create_card("即時摘要")
        right_card.setFixedWidth(360)
        self.summary_record_count = QLabel("資料筆數: 0")
        self.summary_last_sync = QLabel("上次同步: -")
        self.summary_target = QLabel("目標季度: -")
        right_layout.addWidget(self.summary_record_count)
        right_layout.addWidget(self.summary_last_sync)
        right_layout.addWidget(self.summary_target)

        self.latest_table = QTableWidget(0, 4)
        self.latest_table.setHorizontalHeaderLabels(["公司", "最新季度", "Final EPS", "營收"])
        self.latest_table.verticalHeader().setVisible(False)
        self.latest_table.setAlternatingRowColors(True)
        self.latest_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.latest_table.horizontalHeader().setStretchLastSection(True)
        right_layout.addWidget(self.latest_table)

        root_layout.addWidget(left_card)
        root_layout.addWidget(center_card, stretch=1)
        root_layout.addWidget(right_card)
        self.setCentralWidget(root)

        self.select_all_button.clicked.connect(self._on_select_all_companies)
        self.clear_all_button.clicked.connect(self._on_clear_all_companies)
        self.company_list.itemChanged.connect(self._refresh_all_views)

    def _build_collect_tab(self) -> None:
        tab = QWidget()
        tab_layout = QVBoxLayout(tab)
        tab_layout.setContentsMargins(10, 10, 10, 10)
        tab_layout.setSpacing(10)

        controls = QHBoxLayout()
        controls.addWidget(QLabel("目標季度"))
        self.quarter_input = QLineEdit(",".join(DEFAULT_TARGET_QUARTERS))
        self.quarter_input.setPlaceholderText("例如: 2025Q3,2025Q4,2026Q1")
        controls.addWidget(self.quarter_input, stretch=1)

        self.include_latest_checkbox = QCheckBox("加上目前最新季度")
        self.include_latest_checkbox.setChecked(True)
        controls.addWidget(self.include_latest_checkbox)

        self.collect_button = QPushButton("蒐集財報資料")
        self.collect_button.setObjectName("PrimaryButton")
        controls.addWidget(self.collect_button)

        self.export_button = QPushButton("匯出 CSV")
        controls.addWidget(self.export_button)
        tab_layout.addLayout(controls)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        tab_layout.addWidget(self.progress_bar)

        self.log_box = QPlainTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMaximumBlockCount(800)
        self.log_box.setPlaceholderText("執行記錄會顯示在這裡")
        tab_layout.addWidget(self.log_box, stretch=1)

        self.records_table = QTableWidget(0, 9)
        self.records_table.setHorizontalHeaderLabels(
            ["公司", "季度", "營收", "淨利", "EPS(公告)", "EPS(預估)", "PE(可調)", "來源", "更新時間"]
        )
        self.records_table.verticalHeader().setVisible(False)
        self.records_table.setAlternatingRowColors(True)
        self.records_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.records_table.horizontalHeader().setStretchLastSection(True)
        tab_layout.addWidget(self.records_table, stretch=3)

        self.tabs.addTab(tab, "財報蒐集")

        self.collect_button.clicked.connect(self._on_collect_clicked)
        self.export_button.clicked.connect(self._on_export_clicked)

    def _build_eps_tab(self) -> None:
        tab = QWidget()
        tab_layout = QVBoxLayout(tab)
        tab_layout.setContentsMargins(10, 10, 10, 10)
        tab_layout.setSpacing(10)

        top_row = QHBoxLayout()
        top_row.addWidget(QLabel("公司"))
        self.eps_company_combo = QComboBox()
        for profile in self.collector.company_profiles.values():
            self.eps_company_combo.addItem(f"{profile.company_id}  {profile.name}", profile.company_id)
        top_row.addWidget(self.eps_company_combo)

        self.auto_estimate_button = QPushButton("自動預估 EPS")
        self.auto_estimate_button.setObjectName("PrimaryButton")
        top_row.addWidget(self.auto_estimate_button)

        top_row.addStretch(1)
        tab_layout.addLayout(top_row)

        self.eps_table = QTableWidget(0, 6)
        self.eps_table.setHorizontalHeaderLabels(
            ["季度", "EPS(公告)", "EPS(預估)", "Final EPS", "PE(可調)", "目標價"]
        )
        self.eps_table.verticalHeader().setVisible(False)
        self.eps_table.setAlternatingRowColors(True)
        self.eps_table.horizontalHeader().setStretchLastSection(True)
        tab_layout.addWidget(self.eps_table, stretch=2)

        hint = QLabel("操作方式: 點擊表格可輸入 EPS 與 PE，或拖拉圖中的紅色點調整 EPS。")
        hint.setObjectName("HintLabel")
        tab_layout.addWidget(hint)

        self.chart = DraggableEpsCanvas()
        tab_layout.addWidget(self.chart, stretch=3)

        self.tabs.addTab(tab, "EPS 管理")

        self.eps_company_combo.currentIndexChanged.connect(self._on_eps_company_changed)
        self.auto_estimate_button.clicked.connect(self._on_auto_estimate_clicked)
        self.eps_table.cellChanged.connect(self._on_eps_cell_changed)
        self.chart.on_point_changed = self._preview_chart_point
        self.chart.on_drag_finished = self._commit_chart_point

    def _create_card(self, title: str) -> tuple[QFrame, QVBoxLayout]:
        card = QFrame()
        card.setObjectName("Card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)
        title_label = QLabel(title)
        title_label.setObjectName("CardTitle")
        layout.addWidget(title_label)
        return card, layout

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            * {
                font-family: "Noto Sans TC", "Segoe UI";
                font-size: 13px;
            }
            QMainWindow {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 1, y2: 1,
                    stop: 0 #edf3fb,
                    stop: 0.45 #f7fbff,
                    stop: 1 #f3f7ed
                );
            }
            QFrame#Card {
                background-color: rgba(255, 255, 255, 0.95);
                border: 1px solid #d5e0ef;
                border-radius: 15px;
            }
            QLabel#CardTitle {
                font-size: 16px;
                font-weight: 700;
                color: #2b3f5c;
            }
            QLabel#HintLabel {
                color: #44566c;
                font-size: 12px;
            }
            QTabWidget::pane {
                border: 0;
                background: transparent;
            }
            QTabBar::tab {
                background: #eef3ff;
                border: 1px solid #d0dbef;
                padding: 8px 14px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                margin-right: 4px;
            }
            QTabBar::tab:selected {
                background: #dbe9ff;
                color: #1d3557;
                font-weight: 700;
            }
            QPushButton {
                border: 1px solid #b8c9e6;
                border-radius: 8px;
                padding: 6px 12px;
                background: #f5f8ff;
            }
            QPushButton:hover {
                background: #e9f1ff;
            }
            QPushButton#PrimaryButton {
                background: #0f7b6c;
                color: white;
                border: 1px solid #0d6458;
                font-weight: 700;
            }
            QPushButton#PrimaryButton:hover {
                background: #0c675b;
            }
            QLineEdit, QComboBox, QListWidget, QTableWidget, QPlainTextEdit {
                border: 1px solid #ccd8ea;
                border-radius: 8px;
                background: #ffffff;
            }
            QHeaderView::section {
                background: #eff5ff;
                border: 0;
                border-bottom: 1px solid #d6e1f1;
                padding: 6px;
                font-weight: 700;
                color: #2f4d6d;
            }
            """
        )

    def _on_select_all_companies(self) -> None:
        for i in range(self.company_list.count()):
            self.company_list.item(i).setCheckState(Qt.Checked)

    def _on_clear_all_companies(self) -> None:
        for i in range(self.company_list.count()):
            self.company_list.item(i).setCheckState(Qt.Unchecked)

    def _selected_company_ids(self) -> list[str]:
        result: list[str] = []
        for i in range(self.company_list.count()):
            item = self.company_list.item(i)
            if item.checkState() == Qt.Checked:
                result.append(item.data(Qt.UserRole))
        return result

    def _parse_target_quarters(self) -> list[str]:
        text = self.quarter_input.text().strip().replace("，", ",")
        if not text:
            return DEFAULT_TARGET_QUARTERS
        chunks = [item.strip() for item in text.split(",") if item.strip()]
        if not chunks:
            return DEFAULT_TARGET_QUARTERS
        return normalize_quarters(chunks)

    def _append_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_box.appendPlainText(f"[{timestamp}] {message}")

    def _on_collect_clicked(self) -> None:
        company_ids = self._selected_company_ids()
        if not company_ids:
            QMessageBox.warning(self, "提醒", "請至少勾選一家公司。")
            return

        try:
            target_quarters = self._parse_target_quarters()
        except ValueError as exc:
            QMessageBox.warning(self, "季度格式錯誤", str(exc))
            return

        self._append_log(f"開始蒐集: {', '.join(company_ids)}")
        self.progress_bar.setValue(0)
        QApplication.processEvents()

        def on_progress(done: int, total: int, message: str) -> None:
            percent = int(done / total * 100) if total > 0 else 0
            self.progress_bar.setValue(percent)
            self._append_log(message)
            QApplication.processEvents()

        summary = self.collector.collect(
            company_ids=company_ids,
            target_quarters=target_quarters,
            include_latest=self.include_latest_checkbox.isChecked(),
            on_progress=on_progress,
        )

        self._append_log(
            f"完成。成功 {len(summary['success'])} 家，失敗 {len(summary['failed'])} 家。"
        )
        self.progress_bar.setValue(100)
        self._refresh_all_views()

    def _on_export_clicked(self) -> None:
        selected = self._selected_company_ids()
        if not selected:
            QMessageBox.warning(self, "提醒", "請至少勾選一家公司再匯出。")
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        export_path = EXPORT_DIR / f"financial_records_{timestamp}.csv"
        path = self.db.export_csv(export_path=export_path, company_ids=selected)
        self._append_log(f"CSV 匯出完成: {path}")
        QMessageBox.information(self, "匯出完成", f"已匯出到\n{path}")

    def _on_eps_company_changed(self) -> None:
        company_id = self.eps_company_combo.currentData()
        if company_id:
            self._load_eps_company(company_id)

    def _load_eps_company(self, company_id: str) -> None:
        records = self.db.fetch_company_records(company_id)
        if not records:
            profile = self.collector.company_profiles[company_id]
            rows = []
            now_text = datetime.now().isoformat(timespec="seconds")
            for quarter in DEFAULT_TARGET_QUARTERS:
                rows.append(
                    {
                        "company_id": profile.company_id,
                        "company_name": profile.name,
                        "quarter": quarter,
                        "revenue": None,
                        "net_income": None,
                        "eps_reported": None,
                        "eps_estimated": None,
                        "pe_ratio": None,
                        "source": "manual_required",
                        "fetched_at": now_text,
                    }
                )
            self.db.upsert_records(rows)
            records = self.db.fetch_company_records(company_id)

        records.sort(key=lambda item: quarter_sort_key(item["quarter"]))
        self._loading_eps_table = True
        try:
            self.eps_table.setRowCount(len(records))
            quarters: list[str] = []
            chart_values: list[float | None] = []
            editable_mask: list[bool] = []
            for row_idx, record in enumerate(records):
                quarter = record["quarter"]
                reported = _to_float(record.get("eps_reported"))
                estimated = _to_float(record.get("eps_estimated"))
                final_value = reported if reported is not None else estimated
                pe_ratio = _to_float(record.get("pe_ratio"))
                target_price = _calculate_target_price(final_value, pe_ratio)

                quarter_item = QTableWidgetItem(quarter)
                quarter_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                self.eps_table.setItem(row_idx, 0, quarter_item)

                reported_item = QTableWidgetItem(_format_number(reported))
                reported_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                reported_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.eps_table.setItem(row_idx, 1, reported_item)

                estimated_item = QTableWidgetItem(_format_number(estimated))
                estimated_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                if reported is not None:
                    estimated_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                else:
                    estimated_item.setFlags(
                        Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable
                    )
                self.eps_table.setItem(row_idx, 2, estimated_item)

                final_item = QTableWidgetItem(_format_number(final_value))
                final_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                final_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.eps_table.setItem(row_idx, 3, final_item)

                pe_item = QTableWidgetItem(_format_number(pe_ratio, 2))
                pe_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
                pe_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.eps_table.setItem(row_idx, 4, pe_item)

                target_item = QTableWidgetItem(_format_number(target_price, 2))
                target_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                target_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.eps_table.setItem(row_idx, 5, target_item)

                quarters.append(quarter)
                chart_values.append(final_value)
                editable_mask.append(reported is None)
        finally:
            self._loading_eps_table = False

        self.chart.set_series(quarters=quarters, values=chart_values, editable_mask=editable_mask)

    def _on_eps_cell_changed(self, row: int, column: int) -> None:
        if self._loading_eps_table or column not in (2, 4):
            return

        company_id = self.eps_company_combo.currentData()
        if not company_id:
            return

        quarter_item = self.eps_table.item(row, 0)
        value_item = self.eps_table.item(row, column)
        if quarter_item is None or value_item is None:
            return

        quarter = quarter_item.text().strip()
        text = value_item.text().strip()
        if not text or text == "-":
            return

        try:
            numeric_value = float(text.replace(",", ""))
        except ValueError:
            metric = "EPS" if column == 2 else "PE"
            QMessageBox.warning(self, "格式錯誤", f"{metric} 必須是數字。")
            self._load_eps_company(company_id)
            return

        if column == 2:
            self.collector.update_manual_eps(
                company_id=company_id,
                quarter=quarter,
                eps_value=round(numeric_value, 3),
            )
            self._refresh_chart_from_table()
        else:
            self.collector.update_manual_pe(
                company_id=company_id,
                quarter=quarter,
                pe_value=round(numeric_value, 3),
            )

        self._update_final_cell_for_row(row)
        self._update_target_price_cell_for_row(row)
        self._refresh_collection_table()
        self._refresh_summary_panel()

    def _final_eps_from_row(self, row: int) -> float | None:
        reported_item = self.eps_table.item(row, 1)
        estimated_item = self.eps_table.item(row, 2)
        reported = (
            _to_float(reported_item.text().replace(",", ""))
            if reported_item is not None
            else None
        )
        estimated = (
            _to_float(estimated_item.text().replace(",", ""))
            if estimated_item is not None
            else None
        )
        return reported if reported is not None else estimated

    def _update_final_cell_for_row(self, row: int) -> None:
        final_value = self._final_eps_from_row(row)

        final_item = self.eps_table.item(row, 3)
        if final_item is None:
            final_item = QTableWidgetItem()
            final_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            final_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.eps_table.setItem(row, 3, final_item)
        final_item.setText(_format_number(final_value))

    def _update_target_price_cell_for_row(self, row: int) -> None:
        final_eps = self._final_eps_from_row(row)
        pe_item = self.eps_table.item(row, 4)
        pe_ratio = (
            _to_float(pe_item.text().replace(",", ""))
            if pe_item is not None
            else None
        )
        target_price = _calculate_target_price(final_eps, pe_ratio)

        target_item = self.eps_table.item(row, 5)
        if target_item is None:
            target_item = QTableWidgetItem()
            target_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            target_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.eps_table.setItem(row, 5, target_item)
        target_item.setText(_format_number(target_price, 2))

    def _refresh_chart_from_table(self) -> None:
        quarters: list[str] = []
        values: list[float | None] = []
        editable: list[bool] = []
        for row in range(self.eps_table.rowCount()):
            quarter_item = self.eps_table.item(row, 0)
            reported_item = self.eps_table.item(row, 1)
            estimated_item = self.eps_table.item(row, 2)
            if quarter_item is None:
                continue
            quarter = quarter_item.text().strip()
            reported = _to_float(reported_item.text().replace(",", "")) if reported_item else None
            estimated = _to_float(estimated_item.text().replace(",", "")) if estimated_item else None
            final_value = reported if reported is not None else estimated
            quarters.append(quarter)
            values.append(final_value)
            editable.append(reported is None)

        self.chart.set_series(quarters=quarters, values=values, editable_mask=editable)

    def _preview_chart_point(self, quarter: str, value: float) -> None:
        for row in range(self.eps_table.rowCount()):
            quarter_item = self.eps_table.item(row, 0)
            if quarter_item is None or quarter_item.text() != quarter:
                continue
            self._loading_eps_table = True
            try:
                estimated_item = self.eps_table.item(row, 2)
                if estimated_item is None:
                    estimated_item = QTableWidgetItem()
                    estimated_item.setFlags(
                        Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable
                    )
                    estimated_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    self.eps_table.setItem(row, 2, estimated_item)
                estimated_item.setText(_format_number(value))
                self._update_final_cell_for_row(row)
                self._update_target_price_cell_for_row(row)
            finally:
                self._loading_eps_table = False
            break

    def _commit_chart_point(self, quarter: str, value: float) -> None:
        company_id = self.eps_company_combo.currentData()
        if not company_id:
            return
        self.collector.update_manual_eps(company_id=company_id, quarter=quarter, eps_value=value)
        self._append_log(f"{company_id} {quarter} EPS 預估更新為 {value:.3f}")
        self._refresh_collection_table()
        self._refresh_summary_panel()

    def _on_auto_estimate_clicked(self) -> None:
        company_id = self.eps_company_combo.currentData()
        if not company_id:
            return

        existing_rows = self.db.fetch_company_records(company_id)
        known_eps_count = sum(
            1
            for row in existing_rows
            if _to_float(row.get("eps_reported")) is not None
            or _to_float(row.get("eps_estimated")) is not None
        )

        estimates = self.collector.auto_estimate(company_id)
        if estimates:
            message = f"{company_id} 自動預估完成，共更新 {len(estimates)} 季。"
            self._append_log(message)
            QMessageBox.information(self, "自動預估完成", message)
        else:
            if known_eps_count == 0:
                message = (
                    f"{company_id} 目前沒有已知 EPS，請先蒐集財報資料，"
                    "或先手動輸入至少一季 EPS 再按自動預估。"
                )
            else:
                message = f"{company_id} 沒有可預估的季度（可能都已有值）。"
            self._append_log(message)
            QMessageBox.warning(self, "沒有可更新資料", message)

        self._load_eps_company(company_id)
        self._refresh_collection_table()
        self._refresh_summary_panel()

    def _refresh_all_views(self) -> None:
        self._refresh_collection_table()
        self._refresh_summary_panel()
        company_id = self.eps_company_combo.currentData()
        if company_id:
            self._load_eps_company(company_id)

    def _refresh_collection_table(self) -> None:
        selected = self._selected_company_ids()
        rows = self.db.fetch_records(company_ids=selected if selected else None)
        rows.sort(key=lambda item: (item["company_id"], -quarter_sort_key(item["quarter"])))
        self.records_table.setRowCount(len(rows))

        for row_idx, row in enumerate(rows):
            values = [
                f"{row['company_id']} {row['company_name']}",
                row["quarter"],
                _format_money_to_hundred_million(row.get("revenue")),
                _format_money_to_hundred_million(row.get("net_income")),
                _format_number(row.get("eps_reported")),
                _format_number(row.get("eps_estimated")),
                _format_number(row.get("pe_ratio"), 2),
                row.get("source") or "-",
                row.get("fetched_at") or "-",
            ]
            for col_idx, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                if col_idx in (2, 3, 4, 5, 6):
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.records_table.setItem(row_idx, col_idx, item)

    def _refresh_summary_panel(self) -> None:
        selected = self._selected_company_ids()
        record_count = self.db.count_records()
        latest_sync = self.db.latest_sync_time() or "-"
        self.summary_record_count.setText(f"資料筆數: {record_count}")
        self.summary_last_sync.setText(f"上次同步: {latest_sync}")

        try:
            targets = ", ".join(self._parse_target_quarters())
        except ValueError:
            targets = "格式錯誤"
        self.summary_target.setText(f"目標季度: {targets}")

        latest_rows: list[dict[str, Any]] = []
        for company_id in selected:
            company_rows = self.db.fetch_company_records(company_id)
            if not company_rows:
                continue
            company_rows.sort(key=lambda item: quarter_sort_key(item["quarter"]))
            latest_rows.append(company_rows[-1])

        self.latest_table.setRowCount(len(latest_rows))
        for row_idx, row in enumerate(latest_rows):
            final_eps = row.get("eps_reported")
            if final_eps is None:
                final_eps = row.get("eps_estimated")

            display_values = [
                f"{row['company_id']} {row['company_name']}",
                row["quarter"],
                _format_number(final_eps),
                _format_money_to_hundred_million(row.get("revenue")),
            ]
            for col_idx, value in enumerate(display_values):
                item = QTableWidgetItem(str(value))
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                if col_idx in (2, 3):
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.latest_table.setItem(row_idx, col_idx, item)
