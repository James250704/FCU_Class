import sys, io, os, threading, configparser, traceback, pathlib
from pathlib import Path
from contextlib import redirect_stdout, redirect_stderr
import importlib.util
from PySide6.QtGui import QTextCursor
from PySide6.QtCore import Qt, Signal, QObject
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QLineEdit,
    QTextEdit,
    QPushButton,
    QVBoxLayout,
    QHBoxLayout,
    QMessageBox,
    QCheckBox,
    QSpinBox,
)


def import_run_main():
    file_path = pathlib.Path(__file__).with_name("course.py")
    if not file_path.exists():
        raise FileNotFoundError(f"找不到 course.py: {file_path}")
    spec = importlib.util.spec_from_file_location("course", file_path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules["course"] = mod
    spec.loader.exec_module(mod)
    return mod.main


INI = Path("config.ini")
COOKIE_PATTERNS = [
    "cookies*.json",
    "cookies*.txt",
    "cookies*.pkl",
    "cookies*.jar",
    "cookie*.json",
    "cookie*.txt",
    "*.cookie",
    "*.cookies",
    "*.cookiejar",
    "session*_cookies.*",
    "session*.json",
]


class LogEmitter(QObject):
    text = Signal(str)


class QtStream(io.TextIOBase):
    """把 print() 重導到 QTextEdit"""

    def __init__(self, emitter: LogEmitter):
        super().__init__()
        self.emitter = emitter
        self._buf = ""

    def write(self, s):
        if not isinstance(s, str):
            s = s.decode("utf-8", "ignore")
        self._buf += s
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            self.emitter.text.emit(line + "\n")
        return len(s)

    def flush(self):
        if self._buf:
            self.emitter.text.emit(self._buf)
            self._buf = ""


class Runner:
    """在背景執行 course.main()"""

    def __init__(self, append_log):
        self._thread = None
        self._stop_flag = False
        self.append_log = append_log

    def start(self):
        if self._thread and self._thread.is_alive():
            self.append_log("任務已在執行中。\n")
            return
        self._stop_flag = False
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self):
        self._stop_flag = True
        self.append_log(
            "收到停止請求：將在目前請求結束後停止（若網站阻塞可能需等待）。\n"
        )

    def is_stopped(self):
        return self._stop_flag

    def _run(self):
        try:
            run_main = import_run_main()
        except Exception as e:
            self.append_log(
                f"[錯誤] 無法匯入 course.main(): {e}\n{traceback.format_exc()}\n"
            )
            return
        emitter = LogEmitter()
        emitter.text.connect(self.append_log)
        qstream_out = QtStream(emitter)
        qstream_err = QtStream(emitter)
        try:
            with redirect_stdout(qstream_out), redirect_stderr(qstream_err):
                # 傳遞停止檢查函數給 course.main()
                run_main(stop_check_func=self.is_stopped)
        except SystemExit:
            pass
        except Exception as e:
            self.append_log(f"[執行例外] {e}\n{traceback.format_exc()}\n")


class MainWin(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("FCU 選課 GUI")
        self.setMinimumWidth(700)

        # 欄位
        self.ed_nid = QLineEdit()
        self.ed_pwd = QLineEdit()
        self.ed_pwd.setEchoMode(QLineEdit.Password)
        self.ck_show = QCheckBox("顯示密碼")
        self.ck_show.setTristate(False)
        self.ck_show.toggled.connect(self._toggle_pwd)
        self._toggle_pwd(self.ck_show.isChecked())  # 初始化同步

        self.ed_tb = QTextEdit()
        self.ed_tb.setPlaceholderText("課程代號清單：可逗號、空白、換行或 JSON 陣列")
        self._set_two_line_height(self.ed_tb)  # 兩行高度

        # 重試設定
        self.ck_retry = QCheckBox("啟用自動重試")
        self.ck_retry.setTristate(False)
        self.sp_retry_count = QSpinBox()
        self.sp_retry_count.setMinimum(0)  # 0 表示無限重試
        self.sp_retry_count.setMaximum(999)
        self.sp_retry_count.setValue(3)  # 預設重試 3 次
        self.sp_retry_count.setSuffix(" 次")
        self.sp_retry_count.setSpecialValueText("無限重試")  # 當值為 0 時顯示
        self.sp_retry_interval = QSpinBox()
        self.sp_retry_interval.setMinimum(0)  # 0 秒間隔
        self.sp_retry_interval.setMaximum(3600)
        self.sp_retry_interval.setValue(30)  # 預設間隔 30 秒
        self.sp_retry_interval.setSuffix(" 秒")

        # 按鈕
        self.btn_load = QPushButton("讀取 config.ini")
        self.btn_save = QPushButton("儲存 config.ini")
        self.btn_cleancookie = QPushButton("刪除 Cookie")
        self.btn_run = QPushButton("開始執行")
        self.btn_stop = QPushButton("停止")

        # 日誌
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("這裡顯示原本終端列印的流程訊息…")

        # 版面
        top = QVBoxLayout(self)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("學號 NID"))
        row1.addWidget(self.ed_nid)
        top.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("密碼 Password"))
        row2.addWidget(self.ed_pwd)
        row2.addWidget(self.ck_show)
        top.addLayout(row2)

        top.addWidget(QLabel("課程代號(使用,逗號分隔)"))
        top.addWidget(self.ed_tb)

        # 重試設定區域
        retry_layout = QVBoxLayout()
        retry_row1 = QHBoxLayout()
        retry_row1.addWidget(self.ck_retry)
        retry_row1.addStretch()
        retry_layout.addLayout(retry_row1)

        retry_row2 = QHBoxLayout()
        retry_row2.addWidget(QLabel("重試次數"))
        retry_row2.addWidget(self.sp_retry_count)
        retry_row2.addWidget(QLabel("重試間隔"))
        retry_row2.addWidget(self.sp_retry_interval)
        retry_row2.addStretch()
        retry_layout.addLayout(retry_row2)

        top.addLayout(retry_layout)

        row3 = QHBoxLayout()
        row3.addWidget(self.btn_load)
        row3.addWidget(self.btn_save)
        row3.addWidget(self.btn_cleancookie)
        row3.addStretch()
        top.addLayout(row3)

        row4 = QHBoxLayout()
        row4.addWidget(self.btn_run)
        row4.addWidget(self.btn_stop)
        row4.addStretch()
        top.addLayout(row4)

        top.addWidget(QLabel("訊息"))
        top.addWidget(self.log, 1)

        # 行為
        self.btn_load.clicked.connect(self.load_ini)
        self.btn_save.clicked.connect(self.save_ini)
        self.btn_cleancookie.clicked.connect(self.delete_cookies)
        self.btn_run.clicked.connect(self.run_job)
        self.btn_stop.clicked.connect(self.stop_job)

        self.runner = Runner(self.append_log)

        if INI.exists():
            self.load_ini()

    # ---- UI handlers ----
    def _toggle_pwd(self, checked: bool):
        self.ed_pwd.setEchoMode(QLineEdit.Normal if checked else QLineEdit.Password)

    def _set_two_line_height(self, te: QTextEdit):
        fm = te.fontMetrics()
        line_h = fm.lineSpacing()
        padding = 16  # 邊距估計
        te.setFixedHeight(line_h * 2 + padding)

    def append_log(self, s: str):
        cursor = self.log.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.log.setTextCursor(cursor)
        self.log.insertPlainText(s)
        self.log.moveCursor(QTextCursor.End)

    def load_ini(self):
        if not INI.exists():
            QMessageBox.information(
                self, "提示", "目前資料夾找不到 config.ini，請先儲存。"
            )
            return
        cfg = configparser.ConfigParser()
        try:
            cfg.read(INI, encoding="utf-8")
            nid = cfg.get("auth", "NID", fallback="")
            pwd = cfg.get("auth", "PASS", fallback="")
            tb = ""
            if cfg.has_option("course", "tbSubIDs"):
                tb = cfg.get("course", "tbSubIDs", fallback="")
            elif cfg.has_option("course", "tbSubID"):
                tb = cfg.get("course", "tbSubID", fallback="")

            # 讀取重試設定
            retry_enabled = cfg.getboolean("retry", "enabled", fallback=False)
            retry_count = cfg.getint("retry", "count", fallback=3)
            retry_interval = cfg.getint("retry", "interval", fallback=30)

            self.ed_nid.setText(nid)
            self.ed_pwd.setText(pwd)
            self.ed_tb.setPlainText(tb)
            self.ck_retry.setChecked(retry_enabled)
            self.sp_retry_count.setValue(retry_count)
            self.sp_retry_interval.setValue(retry_interval)

            self.append_log("已載入 config.ini。\n")
        except Exception as e:
            QMessageBox.critical(self, "讀取錯誤", str(e))

    def save_ini(self):
        nid = self.ed_nid.text().strip()
        pwd = self.ed_pwd.text().strip()
        tb = self.ed_tb.toPlainText().strip()
        if not nid or not pwd or not tb:
            QMessageBox.warning(self, "缺少欄位", "NID / PASS / tbSubIDs 不可為空。")
            return

        cfg = configparser.ConfigParser()
        cfg["auth"] = {"NID": nid, "PASS": pwd}
        cfg["course"] = {"tbSubIDs": tb}
        cfg["retry"] = {
            "enabled": str(self.ck_retry.isChecked()),
            "count": str(self.sp_retry_count.value()),
            "interval": str(self.sp_retry_interval.value()),
        }

        try:
            with open(INI, "w", encoding="utf-8") as f:
                cfg.write(f)
            self.append_log("已儲存 config.ini。\n")
        except Exception as e:
            QMessageBox.critical(self, "寫入錯誤", str(e))

    def delete_cookies(self):
        cwd = Path.cwd()
        removed = 0
        tried = set()
        for pat in COOKIE_PATTERNS:
            for p in cwd.glob(pat):
                if p in tried or not p.is_file():
                    continue
                tried.add(p)
                try:
                    p.unlink()
                    removed += 1
                    self.append_log(f"已刪除：{p.name}\n")
                except Exception as e:
                    self.append_log(f"無法刪除 {p.name}: {e}\n")
        if removed == 0:
            self.append_log("未找到可刪除的 Cookie 檔案（目前目錄）。\n")
        else:
            self.append_log(f"Cookie 清理完成，共刪除 {removed} 個檔案。\n")

    def run_job(self):
        self.save_ini()
        self.runner.start()
        self.append_log("開始執行。\n")

    def stop_job(self):
        self.runner.stop()


def main():
    app = QApplication(sys.argv)
    w = MainWin()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
