import sys
import os
import subprocess
import platform


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)

from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QFileDialog,
    QHeaderView,
    QAbstractItemView,
    QSplitter,
    QMessageBox,
    QDialog,
    QRadioButton,
    QLineEdit,
    QButtonGroup,
)
from PyQt5.QtCore import Qt, QMimeData
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QFont, QColor, QCursor, QIcon

import database
import config
from converter import ConvertThread


def open_file(path):
    if not os.path.exists(path):
        return False
    if platform.system() == "Windows":
        os.startfile(path)
    elif platform.system() == "Darwin":
        subprocess.run(["open", path])
    else:
        subprocess.run(["xdg-open", path])
    return True


def open_directory(path):
    dir_path = path if os.path.isdir(path) else os.path.dirname(path)
    if not os.path.isdir(dir_path):
        return False
    if platform.system() == "Windows":
        os.startfile(dir_path)
    elif platform.system() == "Darwin":
        subprocess.run(["open", dir_path])
    else:
        subprocess.run(["xdg-open", dir_path])
    return True


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setWindowTitle("MarkItDown GUI - 文件转Markdown工具")
        self.setWindowIcon(QIcon(resource_path("assets/logo.ico")))
        self.setFixedSize(800, 400)

        self.convert_threads = []
        self.current_file = ""
        self.current_output = ""
        self.current_status = ""
        self.current_record_id = None

        self.settings = config.load_config()
        database.init_db()

        self._init_ui()
        self._load_history()
        self._log("应用已启动，请将文件拖放到窗口中进行转换")

    def _init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(4)

        layout.addWidget(self._build_top_panel())

        splitter = QSplitter(Qt.Vertical)
        splitter.addWidget(self._build_history_table())
        splitter.addWidget(self._build_log_area())
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)
        layout.addWidget(splitter)

        self._apply_style()

    def _build_top_panel(self):
        panel = QWidget()
        panel.setObjectName("topPanel")
        hbox = QHBoxLayout(panel)
        hbox.setContentsMargins(6, 4, 6, 4)

        self.lbl_filename = QLabel("选择或拖放文件")
        self.lbl_filename.setObjectName("fileNameLabel")
        self.lbl_filename.setMinimumWidth(180)
        hbox.addWidget(self.lbl_filename)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setMinimumWidth(120)
        hbox.addWidget(self.progress_bar)

        self.lbl_status = QLabel("就绪")
        self.lbl_status.setObjectName("statusLabel")
        self.lbl_status.setMinimumWidth(60)
        hbox.addWidget(self.lbl_status)

        self.btn_select_file = QPushButton("选择文件")
        self.btn_select_file.setFixedSize(60, 26)
        self.btn_select_file.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_select_file.setStyleSheet("""
            QPushButton {
                background-color: #4dabf7;
                border: 1px solid #339af0;
                border-radius: 4px;
                color: #ffffff;
                font-size: 12px;
                padding: 2px 6px;
            }
            QPushButton:hover { background-color: #339af0; }
            QPushButton:pressed { background-color: #228be6; }
        """)
        self.btn_select_file.clicked.connect(self._select_files)
        hbox.insertWidget(0, self.btn_select_file)

        self.btn_settings = QPushButton("\u2699 设置")
        self.btn_settings.setObjectName("settingsBtn")
        self.btn_settings.setFixedSize(64, 26)
        self.btn_settings.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_settings.clicked.connect(self._open_settings)
        hbox.addWidget(self.btn_settings)

        return panel

    def _build_history_table(self):
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(
            ["原始文件", "转换后文件", "结果", "转换时间", "", ""]
        )
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionMode(QAbstractItemView.NoSelection)
        self.table.setFocusPolicy(Qt.NoFocus)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setMouseTracking(True)
        self.table.cellClicked.connect(self._on_table_cell_clicked)
        self.table.cellEntered.connect(self._on_table_cell_entered)
        return self.table

    def _build_log_area(self):
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setObjectName("logArea")
        self.log_text.setPlaceholderText("日志信息...")
        return self.log_text

    def _apply_style(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f8f9fa;
            }
            #topPanel {
                background-color: #ffffff;
                border: 1px solid #dee2e6;
                border-radius: 6px;
            }
            #fileNameLabel {
                color: #495057;
                font-weight: bold;
            }
            #statusLabel {
                color: #868e96;
                font-size: 12px;
            }
            QProgressBar {
                border: 1px solid #dee2e6;
                border-radius: 4px;
                text-align: center;
                background-color: #e9ecef;
                height: 18px;
            }
            QProgressBar::chunk {
                background-color: #4dabf7;
                border-radius: 3px;
            }
            QPushButton {
                background-color: #e9ecef;
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 2px 6px;
                font-size: 12px;
                color: #495057;
            }
            QPushButton:hover {
                background-color: #dee2e6;
            }
            QPushButton:pressed {
                background-color: #ced4da;
            }
            QPushButton:disabled {
                background-color: #f1f3f5;
                color: #adb5bd;
                border-color: #e9ecef;
            }
            #settingsBtn {
                font-size: 12px;
            }
            QTableWidget {
                background-color: #ffffff;
                border: 1px solid #dee2e6;
                border-radius: 6px;
                gridline-color: #e9ecef;
                font-size: 12px;
            }
            QTableWidget::item {
                padding: 2px 4px;
            }
            QHeaderView::section {
                background-color: #f1f3f5;
                border: none;
                border-bottom: 1px solid #dee2e6;
                border-right: 1px solid #e9ecef;
                padding: 4px 6px;
                font-weight: bold;
                font-size: 12px;
                color: #495057;
            }
            #logArea {
                background-color: #212529;
                color: #adb5bd;
                border: 1px solid #dee2e6;
                border-radius: 6px;
                font-family: Consolas, "Courier New", monospace;
                font-size: 11px;
                padding: 4px;
            }
        """)

    def _log(self, message):
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.lbl_status.setText("松开以转换")
            self.lbl_status.setStyleSheet("color: #4dabf7;")

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragLeaveEvent(self, event):
        self.lbl_status.setText(self.current_status or "就绪")
        self.lbl_status.setStyleSheet("")

    def dropEvent(self, event: QDropEvent):
        self.lbl_status.setStyleSheet("")
        urls = event.mimeData().urls()
        if not urls:
            return

        file_paths = []
        for url in urls:
            path = url.toLocalFile()
            if path and os.path.isfile(path):
                file_paths.append(path)

        if not file_paths:
            self._log("未检测到有效文件")
            return

        event.acceptProposedAction()
        for fp in file_paths:
            self._start_convert(fp)

    def _start_convert(self, file_path):
        self.current_file = file_path
        self.current_output = ""
        self.current_status = "正在转换..."
        self.current_record_id = None

        file_name = os.path.basename(file_path)
        self.lbl_filename.setText(file_name)
        self.progress_bar.setValue(0)
        self.lbl_status.setText("正在转换...")

        self._log(f"开始转换: {file_name}")

        output_dir = config.get_output_dir(file_path)
        thread = ConvertThread(file_path, output_dir)
        thread.progress_signal.connect(self._on_progress)
        thread.status_signal.connect(self._on_status)
        thread.log_signal.connect(self._log)
        thread.finished_signal.connect(self._on_convert_finished)
        self.convert_threads.append(thread)
        thread.start()

    def _on_progress(self, value):
        self.progress_bar.setValue(value)

    def _on_status(self, status):
        self.current_status = status
        self.lbl_status.setText(status)
        if status == "转换成功":
            self.lbl_status.setStyleSheet("color: #40c057; font-weight: bold;")
        elif status == "转换失败":
            self.lbl_status.setStyleSheet("color: #fa5252; font-weight: bold;")
        else:
            self.lbl_status.setStyleSheet("color: #4dabf7;")

    def _on_convert_finished(self, source_file, output_file, success, error_msg):
        status = "成功" if success else "失败"
        record_id = database.add_record(source_file, output_file, status, error_msg)

        self.current_output = output_file
        self.current_record_id = record_id
        self.current_status = "转换成功" if success else "转换失败"

        record = database.get_record_by_id(record_id)
        created_at = record["created_at"] if record else None
        self._add_history_row(record_id, source_file, output_file, status, error_msg, created_at=created_at)
        self._log(f"转换完成 [{status}]: {os.path.basename(source_file)}")

    def _select_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "选择待转换的文件", "", "所有文件 (*)"
        )
        if not file_paths:
            return
        for fp in file_paths:
            self._start_convert(fp)

    def _load_history(self):
        records = database.get_all_records()
        self.table.setRowCount(0)
        for rec in records:
            self._add_history_row(
                rec["id"],
                rec["source_file"],
                rec["output_file"],
                rec["status"],
                rec.get("error_message", ""),
                created_at=rec["created_at"],
            )

    def _add_history_row(self, record_id, source_file, output_file, status, error_msg, created_at=None):
        row = self.table.rowCount()
        self.table.insertRow(0)
        self.table.setRowHeight(0, 28)

        source_name = os.path.basename(source_file)
        output_name = os.path.basename(output_file) if output_file else "-"

        item_source = QTableWidgetItem(source_name)
        item_source.setData(Qt.UserRole, source_file)
        item_source.setToolTip(source_file)
        item_source.setForeground(QColor("#4dabf7"))
        src_font = item_source.font()
        src_font.setUnderline(True)
        item_source.setFont(src_font)
        self.table.setItem(0, 0, item_source)

        item_output = QTableWidgetItem(output_name)
        item_output.setData(Qt.UserRole, output_file)
        item_output.setToolTip(output_file)
        item_output.setForeground(QColor("#4dabf7"))
        font = item_output.font()
        font.setUnderline(True)
        item_output.setFont(font)
        self.table.setItem(0, 1, item_output)

        item_status = QTableWidgetItem(status)
        if status == "成功":
            item_status.setForeground(QColor("#40c057"))
        else:
            item_status.setForeground(QColor("#fa5252"))
        if error_msg:
            item_status.setToolTip(error_msg)
        self.table.setItem(0, 2, item_status)

        item_time = QTableWidgetItem(created_at or "")
        self.table.setItem(0, 3, item_time)

        btn_widget = QWidget()
        btn_layout = QHBoxLayout(btn_widget)
        btn_layout.setContentsMargins(2, 0, 2, 0)
        btn_layout.setSpacing(4)

        btn_copy = QPushButton("复制")
        btn_copy.setObjectName("btnCopy")
        btn_copy.setFixedSize(38, 20)
        btn_copy.setCursor(QCursor(Qt.PointingHandCursor))
        btn_copy.setStyleSheet("""
            QPushButton {
                background-color: #748ffc;
                border: 1px solid #5c7cfa;
                border-radius: 4px;
                color: #ffffff;
                font-size: 12px;
                padding: 2px 6px;
            }
            QPushButton:hover { background-color: #5c7cfa; }
            QPushButton:pressed { background-color: #4c6ef5; }
        """)

        def _copy_output(of=output_file):
            if of and os.path.isfile(of):
                with open(of, "r", encoding="utf-8") as f:
                    QApplication.clipboard().setText(f.read())
                self._log(f"已复制文件内容到剪贴板: {os.path.basename(of)}")
            else:
                self._log("输出文件不存在，无法复制")

        btn_copy.clicked.connect(lambda _, of=output_file: _copy_output(of))
        btn_layout.addWidget(btn_copy)

        btn_dir = QPushButton("目录")
        btn_dir.setObjectName("btnDir")
        btn_dir.setFixedSize(38, 20)
        btn_dir.setCursor(QCursor(Qt.PointingHandCursor))
        btn_dir.setStyleSheet("""
            QPushButton {
                background-color: #69db7c;
                border: 1px solid #51cf66;
                border-radius: 4px;
                color: #ffffff;
                font-size: 12px;
                padding: 2px 6px;
            }
            QPushButton:hover { background-color: #51cf66; }
            QPushButton:pressed { background-color: #40c057; }
        """)
        btn_dir.clicked.connect(lambda _, of=output_file: open_directory(of) if of else None)
        btn_layout.addWidget(btn_dir)

        btn_reconvert = QPushButton("重试")
        btn_reconvert.setObjectName("btnReconvert")
        btn_reconvert.setFixedSize(38, 20)
        btn_reconvert.setCursor(QCursor(Qt.PointingHandCursor))
        btn_reconvert.setStyleSheet("""
            QPushButton {
                background-color: #fcc419;
                border: 1px solid #fab005;
                border-radius: 4px;
                color: #ffffff;
                font-size: 12px;
                padding: 2px 6px;
            }
            QPushButton:hover { background-color: #fab005; }
            QPushButton:pressed { background-color: #f59f00; }
        """)
        btn_reconvert.clicked.connect(lambda _, sf=source_file: self._start_convert(sf))
        btn_layout.addWidget(btn_reconvert)

        btn_del = QPushButton("删除")
        btn_del.setObjectName("btnDelete")
        btn_del.setFixedSize(38, 20)
        btn_del.setCursor(QCursor(Qt.PointingHandCursor))
        btn_del.setStyleSheet("""
            QPushButton {
                background-color: #ff6b6b;
                border: 1px solid #fa5252;
                border-radius: 4px;
                color: #ffffff;
                font-size: 12px;
                padding: 2px 6px;
            }
            QPushButton:hover { background-color: #fa5252; }
            QPushButton:pressed { background-color: #f03e3e; }
        """)
        btn_del.clicked.connect(lambda _, rid=record_id: self._delete_history_record(rid))
        btn_layout.addWidget(btn_del)

        self.table.setCellWidget(0, 4, btn_widget)

        item_id = QTableWidgetItem()
        item_id.setData(Qt.UserRole, record_id)
        self.table.setItem(0, 5, item_id)
        self.table.hideColumn(5)

        self.table.setRowHeight(0, 28)

    def _on_table_cell_clicked(self, row, column):
        if column in (0, 1):
            item = self.table.item(row, column)
            if item:
                file_path = item.data(Qt.UserRole)
                if file_path:
                    open_file(file_path)

    def _on_table_cell_entered(self, row, column):
        if column in (0, 1):
            self.table.viewport().setCursor(QCursor(Qt.PointingHandCursor))
        else:
            self.table.viewport().setCursor(QCursor(Qt.ArrowCursor))

    def _update_history_row_output(self, record_id, new_output_file):
        for row in range(self.table.rowCount()):
            id_item = self.table.item(row, 5)
            if id_item and id_item.data(Qt.UserRole) == record_id:
                output_name = os.path.basename(new_output_file) if new_output_file else "-"
                output_item = self.table.item(row, 1)
                if output_item:
                    output_item.setText(output_name)
                    output_item.setData(Qt.UserRole, new_output_file)
                    output_item.setToolTip(new_output_file)
                dir_btn_widget = self.table.cellWidget(row, 4)
                if dir_btn_widget:
                    btn_layout = dir_btn_widget.layout()
                    if btn_layout and btn_layout.count() >= 3:
                        btn_dir = btn_layout.itemAt(2).widget()
                        if btn_dir:
                            btn_dir.clicked.disconnect()
                            btn_dir.clicked.connect(lambda _, of=new_output_file: open_directory(of))
                break

    def _delete_history_record(self, record_id):
        reply = QMessageBox.question(
            self,
            "确认删除",
            "确定要删除该转换记录及输出文件吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            record = database.get_record_by_id(record_id)
            if record:
                if os.path.exists(record["output_file"]):
                    os.remove(record["output_file"])
                    self._log(f"已删除文件: {record['output_file']}")
                database.delete_record(record_id)
                self._remove_history_row(record_id)
                self._log(f"已删除记录: ID={record_id}")

                if self.current_record_id == record_id:
                    self.current_file = ""
                    self.current_output = ""
                    self.current_status = "就绪"
                    self.current_record_id = None
                    self.lbl_filename.setText("选择或拖放文件")
                    self.progress_bar.setValue(0)
                    self.lbl_status.setText("就绪")
                    self.lbl_status.setStyleSheet("")

    def _remove_history_row(self, record_id):
        for row in range(self.table.rowCount()):
            id_item = self.table.item(row, 5)
            if id_item and id_item.data(Qt.UserRole) == record_id:
                self.table.removeRow(row)
                break

    def _open_settings(self):
        dialog = SettingsDialog(self.settings, self)
        if dialog.exec_() == QDialog.Accepted:
            self.settings = config.load_config()
            self._log("设置已更新")


class SettingsDialog(QDialog):
    def __init__(self, current_settings, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.setFixedSize(460, 220)
        self.setModal(True)
        self._settings = dict(current_settings)
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 12)
        layout.setSpacing(12)

        dir_label = QLabel("转换后文件输出目录：")
        dir_label.setStyleSheet("font-weight: bold; font-size: 13px; color: #343a40;")
        layout.addWidget(dir_label)

        self.btn_group = QButtonGroup(self)
        self.radio_original = QRadioButton("原始文件目录（默认）")
        self.radio_original.setObjectName("radioOption")
        self.btn_group.addButton(self.radio_original, 0)
        layout.addWidget(self.radio_original)

        custom_row = QHBoxLayout()
        self.radio_custom = QRadioButton("自定义目录：")
        self.radio_custom.setObjectName("radioOption")
        self.btn_group.addButton(self.radio_custom, 1)
        custom_row.addWidget(self.radio_custom)

        self.txt_dir = QLineEdit()
        self.txt_dir.setPlaceholderText("选择或输入自定义输出目录路径...")
        self.txt_dir.setMinimumWidth(200)
        self.txt_dir.setEnabled(False)
        custom_row.addWidget(self.txt_dir)

        self.btn_browse = QPushButton("浏览")
        self.btn_browse.setFixedSize(56, 26)
        self.btn_browse.setEnabled(False)
        self.btn_browse.clicked.connect(self._browse_dir)
        custom_row.addWidget(self.btn_browse)
        layout.addLayout(custom_row)

        self.btn_group.buttonClicked.connect(self._on_radio_changed)

        self._load_settings()

        layout.addStretch()

        btn_row = QHBoxLayout()
        btn_row.addStretch()

        self.btn_ok = QPushButton("确定")
        self.btn_ok.setFixedSize(72, 30)
        self.btn_ok.setObjectName("confirmBtn")
        self.btn_ok.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_ok.clicked.connect(self._on_ok)
        btn_row.addWidget(self.btn_ok)

        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.setFixedSize(72, 30)
        self.btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(self.btn_cancel)
        layout.addLayout(btn_row)

        self._apply_style()

    def _load_settings(self):
        mode = self._settings.get("output_mode", "original")
        custom_dir = self._settings.get("custom_output_dir", "")

        if mode == "custom":
            self.radio_custom.setChecked(True)
            self.txt_dir.setText(custom_dir)
            self.txt_dir.setEnabled(True)
            self.btn_browse.setEnabled(True)
        else:
            self.radio_original.setChecked(True)
            self.txt_dir.setEnabled(False)
            self.btn_browse.setEnabled(False)

    def _on_radio_changed(self, btn):
        is_custom = (btn == self.radio_custom)
        self.txt_dir.setEnabled(is_custom)
        self.btn_browse.setEnabled(is_custom)

    def _browse_dir(self):
        current = self.txt_dir.text().strip()
        if not current or not os.path.isdir(current):
            current = config._get_default_output_dir()
        selected = QFileDialog.getExistingDirectory(
            self, "选择输出目录", current
        )
        if selected:
            self.txt_dir.setText(selected)

    def _on_ok(self):
        if self.radio_custom.isChecked():
            custom_dir = self.txt_dir.text().strip()
            if not custom_dir:
                QMessageBox.warning(self, "提示", "请输入或选择自定义输出目录路径。")
                return
            valid, msg = config.validate_output_dir(custom_dir)
            if not valid:
                QMessageBox.warning(self, "目录无效", msg)
                return
            self._settings["output_mode"] = "custom"
            self._settings["custom_output_dir"] = custom_dir
        else:
            self._settings["output_mode"] = "original"

        config.save_config(self._settings)
        self.accept()

    def _apply_style(self):
        self.setStyleSheet("""
            QDialog {
                background-color: #f8f9fa;
            }
            #radioOption {
                font-size: 12px;
                color: #495057;
                spacing: 6px;
            }
            QLineEdit {
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 4px 8px;
                font-size: 12px;
                background-color: #ffffff;
                color: #495057;
            }
            QLineEdit:disabled {
                background-color: #f1f3f5;
                color: #adb5bd;
            }
            QPushButton {
                background-color: #e9ecef;
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 2px 6px;
                font-size: 12px;
                color: #495057;
            }
            QPushButton:hover {
                background-color: #dee2e6;
            }
            QPushButton:pressed {
                background-color: #ced4da;
            }
            QPushButton:disabled {
                background-color: #f1f3f5;
                color: #adb5bd;
                border-color: #e9ecef;
            }
            #confirmBtn {
                background-color: #4dabf7;
                color: #ffffff;
                border-color: #339af0;
            }
            #confirmBtn:hover {
                background-color: #339af0;
            }
            #confirmBtn:pressed {
                background-color: #228be6;
            }
        """)


def main():
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei", 9))
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
