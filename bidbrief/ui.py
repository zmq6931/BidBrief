"""PySide6 desktop UI: folder picker, extension checkboxes, progress and stop."""
import os
import threading
from datetime import datetime

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QFileDialog, QGridLayout, QGroupBox, QHBoxLayout,
    QLabel, QLineEdit, QMainWindow, QMessageBox, QPlainTextEdit, QProgressBar,
    QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

from . import config as config_mod
from .pipeline import run_and_export
from .scanner import SUPPORTED_EXTENSIONS

COL_FILE, COL_STATUS, COL_PAGES, COL_ITEMS = 0, 1, 2, 3


class Worker(QThread):
    sig_files = Signal(list)                      # file paths, to populate the table
    sig_log = Signal(str)
    sig_file = Signal(int, str, object, int)      # row, status, page count, item count
    sig_progress = Signal(int, int)               # files done, total
    sig_done = Signal(dict)

    def __init__(self, cfg, folder, exts, recursive, out_dir, formats):
        super().__init__()
        self.cfg = cfg
        self.folder = folder
        self.exts = exts
        self.recursive = recursive
        self.out_dir = out_dir
        self.formats = formats
        self._stop = threading.Event()

    def stop(self):
        self._stop.set()

    def run(self):
        from .scanner import scan_folder

        files = scan_folder(self.folder, self.exts, self.recursive)
        if not files:
            self.sig_log.emit("所选文件夹中没有匹配的文件。")
            self.sig_done.emit({"cancelled": False, "exports": {}})
            return
        self.sig_files.emit(files)
        self.sig_log.emit(f"共找到 {len(files)} 个文件，开始处理。")

        def on_file(row, status, pages, items):
            self.sig_file.emit(row, status, pages, items)

        def on_chunk(row, ci, total):
            pass  # progress already surfaced through on_file statuses

        try:
            result = run_and_export(
                self.cfg, files, self.out_dir, cancel=self._stop,
                formats=self.formats,
                on_log=self.sig_log.emit, on_file=on_file, on_chunk=on_chunk,
            )
            self.sig_done.emit(result)
        except Exception as e:
            self.sig_done.emit({"error": str(e), "exports": {}})


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("BidBrief — 招标文件要点提取")
        self.resize(1000, 780)
        self.worker = None
        self.last_out_dir = None

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        root.addWidget(self._build_folder_group())
        root.addWidget(self._build_settings_group())
        root.addWidget(self._build_prompt_group())
        root.addWidget(self._build_run_group())
        root.addWidget(self._build_table())
        root.addWidget(self._build_log())

    # ------------------------------------------------------------------
    def _build_folder_group(self):
        grp = QGroupBox("一、选择文件")
        lay = QGridLayout(grp)

        lay.addWidget(QLabel("文件夹："), 0, 0)
        self.ed_folder = QLineEdit()
        lay.addWidget(self.ed_folder, 0, 1)
        btn = QPushButton("浏览…")
        btn.clicked.connect(self._pick_folder)
        lay.addWidget(btn, 0, 2)

        lay.addWidget(QLabel("文件类型："), 1, 0)
        ext_lay = QHBoxLayout()
        self.ext_boxes = {}
        for ext in SUPPORTED_EXTENSIONS:
            cb = QCheckBox(ext)
            cb.setChecked(ext == ".pdf")
            self.ext_boxes[ext] = cb
            ext_lay.addWidget(cb)
        self.cb_recursive = QCheckBox("包含子文件夹")
        self.cb_recursive.setChecked(True)
        ext_lay.addWidget(self.cb_recursive)
        ext_lay.addStretch()
        lay.addLayout(ext_lay, 1, 1, 1, 2)

        lay.addWidget(QLabel("输出目录："), 2, 0)
        self.ed_out = QLineEdit()
        self.ed_out.setPlaceholderText("留空则保存到所选文件夹下的 BidBrief_输出")
        lay.addWidget(self.ed_out, 2, 1)
        btn2 = QPushButton("浏览…")
        btn2.clicked.connect(self._pick_out)
        lay.addWidget(btn2, 2, 2)

        lay.addWidget(QLabel("导出格式："), 3, 0)
        fmt_lay = QHBoxLayout()
        self.fmt_boxes = {}
        for key, label in (("xlsx", "Excel(.xlsx)"), ("docx", "Word(.docx)"),
                           ("pptx", "PPT(.pptx)"), ("md", "Markdown(.md)")):
            cb = QCheckBox(label)
            cb.setChecked(True)
            self.fmt_boxes[key] = cb
            fmt_lay.addWidget(cb)
        fmt_lay.addStretch()
        lay.addLayout(fmt_lay, 3, 1, 1, 2)
        return grp

    def _build_settings_group(self):
        grp = QGroupBox("二、模型设置（DeepSeek / OpenAI 兼容接口）")
        lay = QGridLayout(grp)
        lay.addWidget(QLabel("API Key："), 0, 0)
        self.ed_key = QLineEdit()
        self.ed_key.setEchoMode(QLineEdit.EchoMode.Password)
        cb_show = QCheckBox("显示")
        cb_show.toggled.connect(
            lambda on: self.ed_key.setEchoMode(
                QLineEdit.EchoMode.Normal if on else QLineEdit.EchoMode.Password))
        lay.addWidget(self.ed_key, 0, 1)
        lay.addWidget(cb_show, 0, 2)

        lay.addWidget(QLabel("接口地址："), 1, 0)
        self.ed_base = QLineEdit()
        lay.addWidget(self.ed_base, 1, 1)
        lay.addWidget(QLabel("模型："), 2, 0)
        self.ed_model = QLineEdit()
        lay.addWidget(self.ed_model, 2, 1)

        cfg = config_mod.load_config()
        self.ed_key.setText(cfg.get("api_key", ""))
        self.ed_base.setText(cfg.get("base_url", ""))
        self.ed_model.setText(cfg.get("model", ""))
        return grp

    def _build_prompt_group(self):
        grp = QGroupBox("三、自定义抽取要求（可选）")
        lay = QVBoxLayout(grp)
        self.ed_prompt = QPlainTextEdit()
        self.ed_prompt.setMaximumHeight(90)
        self.ed_prompt.setPlaceholderText(
            "留空则按默认六类（资质/商务/技术/时间节点/废标项/评分标准）全量抽取。\n"
            "例如：本次仅做劳务分包，请重点抽取分包资质要求、总包配合费、"
            "禁止分包/转包条款、与分包相关的废标项，其余内容可忽略。"
        )
        lay.addWidget(self.ed_prompt)
        hint = QLabel(
            "说明：文档按约5000字分块依次送入模型，单个文件大小不限，"
            "处理时间随页数增长（约每块30~60秒）。"
        )
        hint.setStyleSheet("color:#666;")
        lay.addWidget(hint)
        cfg = config_mod.load_config()
        self.ed_prompt.setPlainText(cfg.get("custom_prompt", ""))
        return grp

    def _build_run_group(self):
        grp = QGroupBox("四、运行")
        lay = QHBoxLayout(grp)
        self.btn_start = QPushButton("开始")
        self.btn_start.setMinimumHeight(34)
        f = QFont(); f.setBold(True)
        self.btn_start.setFont(f)
        self.btn_start.setStyleSheet("QPushButton{background:#2e7d32;color:white;border-radius:4px}"
                                     "QPushButton:disabled{background:#9e9e9e}")
        self.btn_start.clicked.connect(self._start)
        self.btn_stop = QPushButton("停止")
        self.btn_stop.setMinimumHeight(34)
        self.btn_stop.setEnabled(False)
        self.btn_stop.setStyleSheet("QPushButton{background:#c62828;color:white;border-radius:4px}"
                                    "QPushButton:disabled{background:#9e9e9e}")
        self.btn_stop.clicked.connect(self._stop)
        self.btn_open = QPushButton("打开输出目录")
        self.btn_open.setEnabled(False)
        self.btn_open.clicked.connect(self._open_out)
        self.progress = QProgressBar()
        self.progress.setFixedWidth(260)
        self.lb_progress = QLabel("就绪")
        lay.addWidget(self.btn_start)
        lay.addWidget(self.btn_stop)
        lay.addWidget(self.btn_open)
        lay.addStretch()
        lay.addWidget(self.lb_progress)
        lay.addWidget(self.progress)
        return grp

    def _build_table(self):
        grp = QGroupBox("文件列表")
        lay = QVBoxLayout(grp)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["文件", "状态", "页数", "要点数"])
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.setColumnWidth(COL_FILE, 430)
        self.table.setColumnWidth(COL_STATUS, 220)
        self.table.setColumnWidth(COL_PAGES, 60)
        self.table.setColumnWidth(COL_ITEMS, 60)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        lay.addWidget(self.table)
        return grp

    def _build_log(self):
        grp = QGroupBox("日志")
        lay = QVBoxLayout(grp)
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(110)
        lay.addWidget(self.log)
        return grp

    # ------------------------------------------------------------------
    def _pick_folder(self):
        d = QFileDialog.getExistingDirectory(self, "选择招标文件所在文件夹")
        if d:
            self.ed_folder.setText(d)

    def _pick_out(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if d:
            self.ed_out.setText(d)

    def _log(self, msg):
        self.log.appendPlainText(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

    # ------------------------------------------------------------------
    def _start(self):
        folder = self.ed_folder.text().strip()
        if not folder or not os.path.isdir(folder):
            QMessageBox.warning(self, "提示", "请先选择有效的文件夹。")
            return
        exts = [ext for ext, cb in self.ext_boxes.items() if cb.isChecked()]
        if not exts:
            QMessageBox.warning(self, "提示", "请至少勾选一种文件类型。")
            return
        formats = [key for key, cb in self.fmt_boxes.items() if cb.isChecked()]
        if not formats:
            QMessageBox.warning(self, "提示", "请至少勾选一种导出格式。")
            return

        cfg = config_mod.load_config()
        cfg["api_key"] = self.ed_key.text().strip()
        cfg["base_url"] = self.ed_base.text().strip() or cfg["base_url"]
        cfg["model"] = self.ed_model.text().strip() or cfg["model"]
        cfg["custom_prompt"] = self.ed_prompt.toPlainText().strip()
        if not cfg["api_key"]:
            QMessageBox.warning(self, "提示", "请填写 API Key。")
            return
        config_mod.save_config({
            "api_key": cfg["api_key"], "base_url": cfg["base_url"],
            "model": cfg["model"], "custom_prompt": cfg["custom_prompt"],
        })

        out_dir = self.ed_out.text().strip() or os.path.join(folder, "BidBrief_输出")
        os.makedirs(out_dir, exist_ok=True)

        self.table.setRowCount(0)
        self.log.clear()
        self.progress.setRange(0, 0)  # indeterminate until the file list arrives
        self.lb_progress.setText("扫描中…")
        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.btn_open.setEnabled(False)
        self._log(f"输出目录：{out_dir}")

        self.worker = Worker(cfg, folder, exts, self.cb_recursive.isChecked(),
                             out_dir, formats)
        self.worker.sig_files.connect(self._on_files)
        self.worker.sig_file.connect(self._on_file_status)
        self.worker.sig_log.connect(self._log)
        self.worker.sig_done.connect(self._on_done)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.start()

    def _stop(self):
        if self.worker is not None:
            self.worker.stop()
            self.btn_stop.setEnabled(False)
            self.lb_progress.setText("正在停止（等待当前请求结束）…")
            self._log("已请求停止：当前小块完成后中止，已完成文件将照常导出。")

    def _open_out(self):
        if self.last_out_dir and os.path.isdir(self.last_out_dir):
            os.startfile(self.last_out_dir)

    # ------------------------------------------------------------------
    def _on_files(self, files):
        self.progress.setRange(0, len(files))
        self.progress.setValue(0)
        self.table.setRowCount(len(files))
        for i, path in enumerate(files):
            it = QTableWidgetItem(os.path.basename(path))
            it.setToolTip(path)
            self.table.setItem(i, COL_FILE, it)
            self.table.setItem(i, COL_STATUS, QTableWidgetItem("待处理"))
            self.table.setItem(i, COL_PAGES, QTableWidgetItem("—"))
            self.table.setItem(i, COL_ITEMS, QTableWidgetItem("—"))

    def _on_file_status(self, row, status, pages, items):
        if 0 <= row < self.table.rowCount():
            if self.table.item(row, COL_STATUS):
                self.table.item(row, COL_STATUS).setText(status)
            if pages is not None:
                self.table.item(row, COL_PAGES).setText(str(pages))
            if items is not None:
                self.table.item(row, COL_ITEMS).setText(str(items))
            done = status.startswith(("完成", "失败", "跳过", "已停止"))
            if done:
                self.progress.setValue(self.progress.value() + 1)
                self.lb_progress.setText(f"{self.progress.value()}/{self.progress.maximum()}")

    def _on_done(self, result):
        self.btn_start.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.progress.setRange(0, 1)
        self.progress.setValue(1)
        frs = result.get("file_results", [])
        n_ok = sum(1 for f in frs if f.get("status", "").startswith("完成"))
        n_skip = sum(1 for f in frs if f.get("status", "").startswith("跳过"))
        n_fail = sum(1 for f in frs if f.get("status", "").startswith("失败"))
        n_stop = sum(1 for f in frs if f.get("status", "").startswith("已停止"))
        cancelled = result.get("cancelled", False)

        exports = result.get("exports", {})
        if exports:
            self.last_out_dir = os.path.dirname(
                exports.get("excel") or exports.get("word")
                or exports.get("ppt") or exports.get("markdown") or "")
            self.btn_open.setEnabled(True)

        if "error" in result:
            QMessageBox.critical(self, "错误", f"运行出错：{result['error']}")
        else:
            msg = (f"处理完成：成功 {n_ok}，跳过 {n_skip}，失败 {n_fail}，中止 {n_stop}。"
                   + ("（用户停止）" if cancelled else ""))
            for k, v in exports.items():
                if k != "raw_dir":
                    msg += f"\n{k}: {os.path.basename(v)}"
            if exports:
                msg += f"\n\n保存位置：{self.last_out_dir}"
            QMessageBox.information(self, "完成", msg)
            self.lb_progress.setText(msg.split("\n")[0])
            self._log(f"结束。成功 {n_ok} / 跳过 {n_skip} / 失败 {n_fail} / 中止 {n_stop}")


def run_app():
    app = QApplication([])
    win = MainWindow()
    win.show()
    app.exec()
