# -*- coding: utf-8 -*-
# Oracle 批量模糊查询小工具（tkinter + python-oracledb + pandas/openpyxl）
#
# 依赖安装：
#     pip install oracledb pandas openpyxl
#
# 启动：
#     python oracle_batch_query.py
#
# 打包 exe：
#     双击 build_exe.bat，或运行 build_exe.ps1

import base64
import csv
import ctypes
import datetime
import io
import json
import os
import re
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import oracledb
import pandas as pd
from openpyxl.utils import get_column_letter

APP_TITLE = "Oracle 批量模糊查询工具"
BATCH_SIZE = 200
MAX_DISPLAY_ROWS = 2000
IDENT_RE = re.compile(r"^[A-Za-z0-9_]+(?:\.[A-Za-z0-9_]+)*$")
BAD_EXCEL_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def normalize_value(value):
    if value is None:
        return None
    if hasattr(oracledb, "LOB") and isinstance(value, oracledb.LOB):
        try:
            value = value.read()
        except Exception:
            value = str(value)
    if isinstance(value, (bytes, bytearray, memoryview)):
        return bytes(value).hex()
    if isinstance(value, datetime.datetime) and value.tzinfo is not None:
        return value.strftime("%Y-%m-%d %H:%M:%S %z")
    if isinstance(value, datetime.time) and value.tzinfo is not None:
        return value.strftime("%H:%M:%S %z")
    if isinstance(value, datetime.timedelta):
        return str(value)
    return value


def freeze_value(value):
    try:
        hash(value)
    except TypeError:
        return (type(value).__name__, str(value))
    return (type(value).__name__, value)


def cell_text(value):
    if value is None:
        return ""
    text = str(value)
    if len(text) > 500:
        return text[:497] + "..."
    return text


def like_regex(value):
    parts = []
    for char in f"%{value}%":
        if char == "%":
            parts.append(".*")
        elif char == "_":
            parts.append(".")
        else:
            parts.append(re.escape(char))
    return re.compile("^" + "".join(parts) + "$", re.IGNORECASE | re.DOTALL)


def column_key(name):
    return str(name).strip().split(".")[-1].lower()


def excel_safe(value):
    if isinstance(value, str):
        return BAD_EXCEL_CHARS_RE.sub("", value)
    return value


def autofit_sheet(worksheet):
    for column_cells in worksheet.columns:
        letter = get_column_letter(column_cells[0].column)
        max_width = 8
        for index, cell in enumerate(column_cells):
            if index >= 300:
                break
            if cell.value is None:
                continue
            width = sum(2 if ord(char) > 127 else 1 for char in str(cell.value))
            if width > max_width:
                max_width = width
        worksheet.column_dimensions[letter].width = min(max_width + 2, 50)


def friendly_db_error(exc):
    text = str(exc).strip()
    rules = [
        ("ORA-01017", "用户名或密码错误，请检查登录凭据。"),
        ("ORA-12541", "数据库监听未启动或地址错误（无监听程序）。"),
        ("ORA-12514", "连接串中的服务名不存在，请检查 service_name。"),
        ("ORA-12170", "连接数据库超时，请检查网络、防火墙和连接串。"),
        ("ORA-00942", "表或视图不存在，请检查表名以及当前用户是否有权限。"),
        ("ORA-00904", "字段名无效，请检查匹配字段和查询字段的拼写。"),
        ("ORA-00933", "SQL 语法错误，请检查表名和字段名。"),
        ("ORA-00936", "SQL 语句缺少表达式，请检查查询字段。"),
        ("ORA-01722", "数值转换失败，请检查字段类型。"),
        ("ORA-01861", "日期格式不匹配，请检查字段类型。"),
        ("DPY-6005", "无法连接数据库，请检查连接串、网络和数据库状态。"),
        ("DPY-4011", "数据库连接已断开或被服务端关闭。"),
        ("DPY-6000", "连接串格式无法解析，请使用 host:port/service_name 或 TNS 别名。"),
    ]
    for code, message in rules:
        if code in text:
            return f"{message}\n\n详细信息：\n{text}"
    return f"操作失败：\n{text}"


CONFIG_PATH = (
    Path(os.environ.get("APPDATA") or Path.home())
    / "OracleBatchQuery"
    / "connection.json"
)


def _build_dpapi():
    from ctypes import wintypes

    class DataBlob(ctypes.Structure):
        _fields_ = [
            ("cbData", wintypes.DWORD),
            ("pbData", ctypes.POINTER(ctypes.c_char)),
        ]

    crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    crypt32.CryptProtectData.argtypes = [
        ctypes.POINTER(DataBlob),
        wintypes.LPCWSTR,
        ctypes.POINTER(DataBlob),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(DataBlob),
    ]
    crypt32.CryptProtectData.restype = wintypes.BOOL

    crypt32.CryptUnprotectData.argtypes = [
        ctypes.POINTER(DataBlob),
        ctypes.POINTER(wintypes.LPWSTR),
        ctypes.POINTER(DataBlob),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(DataBlob),
    ]
    crypt32.CryptUnprotectData.restype = wintypes.BOOL

    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p

    return {
        "DataBlob": DataBlob,
        "crypt32": crypt32,
        "kernel32": kernel32,
    }


try:
    _DPAPI = _build_dpapi() if sys.platform == "win32" else None
except Exception:
    _DPAPI = None


def _make_dpapi_blob(data):
    data_blob = _DPAPI["DataBlob"]
    buffer = ctypes.create_string_buffer(data)
    return data_blob(
        len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_char))
    ), buffer


def protect_password(password):
    if not password:
        return ""
    if _DPAPI is None:
        raise RuntimeError("当前系统不支持本机加密保存密码（仅 Windows 可用）。")
    data = password.encode("utf-8")
    in_blob, _buffer = _make_dpapi_blob(data)
    out_blob = _DPAPI["DataBlob"]()
    if not _DPAPI["crypt32"].CryptProtectData(
        ctypes.byref(in_blob),
        "OracleBatchQuery",
        None,
        None,
        None,
        0x1,
        ctypes.byref(out_blob),
    ):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        encrypted = ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        if out_blob.pbData:
            _DPAPI["kernel32"].LocalFree(
                ctypes.cast(out_blob.pbData, ctypes.c_void_p)
            )
    return base64.b64encode(encrypted).decode("ascii")


def unprotect_password(token):
    if not token:
        return ""
    if _DPAPI is None:
        raise RuntimeError("当前系统不支持本机加密保存密码（仅 Windows 可用）。")
    try:
        encrypted = base64.b64decode(token, validate=True)
    except Exception as exc:
        raise ValueError("保存的密码数据格式不正确。") from exc
    in_blob, _buffer = _make_dpapi_blob(encrypted)
    out_blob = _DPAPI["DataBlob"]()
    if not _DPAPI["crypt32"].CryptUnprotectData(
        ctypes.byref(in_blob),
        None,
        None,
        None,
        None,
        0x1,
        ctypes.byref(out_blob),
    ):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        data = ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        if out_blob.pbData:
            _DPAPI["kernel32"].LocalFree(
                ctypes.cast(out_blob.pbData, ctypes.c_void_p)
            )
    return data.decode("utf-8")


def load_connection_config():
    if not CONFIG_PATH.is_file():
        return {}
    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def save_connection_config(user, dsn, password, remember):
    payload = {
        "user": user,
        "dsn": dsn,
        "remember": bool(remember),
        "password_enc": "",
    }
    if remember and password:
        payload["password_enc"] = protect_password(password)
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    temp_path = CONFIG_PATH.with_name(CONFIG_PATH.name + ".tmp")
    temp_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    os.replace(temp_path, CONFIG_PATH)


def clear_connection_config():
    try:
        CONFIG_PATH.unlink()
    except FileNotFoundError:
        pass


def parse_clipboard_column(text):
    normalized = text.replace("\r\n", "\n").replace("\r", "\n").lstrip("\ufeff")
    rows = list(csv.reader(io.StringIO(normalized), delimiter="\t"))
    multiple_columns = any(len(row) > 1 for row in rows)
    values = []
    for row in rows:
        if not row:
            continue
        value = row[0].strip()
        if value:
            values.append(value)
    return values, multiple_columns


class OracleBatchQueryApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1180x800")
        self.root.minsize(920, 640)

        self.user_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.dsn_var = tk.StringVar()
        self.table_var = tk.StringVar()
        self.match_var = tk.StringVar()
        self.select_fields_var = tk.StringVar()
        self.remember_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="就绪：请填写连接信息和查询条件")

        self._busy = False
        self._result_ready = False
        self._result_columns = []
        self._result_rows = []
        self._unmatched_values = []
        self._status_summary = ""
        self._render_token = 0

        self._build_ui()
        self._load_saved_connection()
        self._set_busy(False)

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(4, weight=1)

        connection_frame = ttk.LabelFrame(main, text="数据库连接", padding=8)
        connection_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        connection_frame.columnconfigure(5, weight=1)

        ttk.Label(connection_frame, text="用户名").grid(row=0, column=0, sticky="w", padx=(0, 4))
        self.user_entry = ttk.Entry(connection_frame, textvariable=self.user_var, width=18)
        self.user_entry.grid(row=0, column=1, sticky="ew", padx=(0, 12))
        ttk.Label(connection_frame, text="密码").grid(row=0, column=2, sticky="w", padx=(0, 4))
        self.password_entry = ttk.Entry(
            connection_frame, textvariable=self.password_var, show="*", width=18
        )
        self.password_entry.grid(row=0, column=3, sticky="ew", padx=(0, 12))
        ttk.Label(connection_frame, text="连接串").grid(row=0, column=4, sticky="w", padx=(0, 4))
        self.dsn_entry = ttk.Entry(connection_frame, textvariable=self.dsn_var)
        self.dsn_entry.grid(row=0, column=5, sticky="ew", padx=(0, 4))
        ttk.Label(connection_frame, text="如 host:1521/service_name", foreground="#666666").grid(
            row=1, column=5, sticky="w"
        )
        self.test_btn = ttk.Button(connection_frame, text="测试连接", command=self._test_connection)
        self.test_btn.grid(row=0, column=6, padx=(8, 0), sticky="ew")
        self.save_conn_btn = ttk.Button(
            connection_frame, text="保存连接", command=self._save_connection
        )
        self.save_conn_btn.grid(row=1, column=6, padx=(8, 0), pady=(2, 0), sticky="ew")
        self.clear_saved_btn = ttk.Button(
            connection_frame, text="清除保存", command=self._clear_saved_connection
        )
        self.clear_saved_btn.grid(row=2, column=6, padx=(8, 0), pady=(6, 0), sticky="ew")

        self.remember_check = ttk.Checkbutton(
            connection_frame,
            text="记住密码（Windows 本机加密保存）",
            variable=self.remember_var,
        )
        self.remember_check.grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(6, 0)
        )
        ttk.Label(
            connection_frame,
            text="勾选后点“保存连接”立即保存；查询或测试成功后也会自动保存",
            foreground="#666666",
        ).grid(row=2, column=4, columnspan=2, sticky="w", pady=(6, 0))

        query_frame = ttk.LabelFrame(main, text="查询设置", padding=8)
        query_frame.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        query_frame.columnconfigure(1, weight=2)
        query_frame.columnconfigure(3, weight=1)
        query_frame.columnconfigure(5, weight=2)

        ttk.Label(query_frame, text="表名").grid(row=0, column=0, sticky="w", padx=(0, 4))
        ttk.Entry(query_frame, textvariable=self.table_var).grid(
            row=0, column=1, sticky="ew", padx=(0, 12)
        )
        ttk.Label(query_frame, text="匹配字段").grid(row=0, column=2, sticky="w", padx=(0, 4))
        ttk.Entry(query_frame, textvariable=self.match_var).grid(
            row=0, column=3, sticky="ew", padx=(0, 12)
        )
        ttk.Label(query_frame, text="查询字段").grid(row=0, column=4, sticky="w", padx=(0, 4))
        ttk.Entry(query_frame, textvariable=self.select_fields_var).grid(
            row=0, column=5, sticky="ew"
        )
        ttk.Label(
            query_frame,
            text="查询字段用逗号分隔，留空表示 SELECT *；表名和字段名仅允许字母、数字、下划线和点号",
            foreground="#666666",
        ).grid(row=1, column=5, columnspan=2, sticky="w", pady=(4, 0))

        values_frame = ttk.LabelFrame(
            main,
            text="待匹配值（每行一个，空行和重复值会自动忽略）",
            padding=8,
        )
        values_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 8))
        values_frame.columnconfigure(0, weight=1)
        values_frame.rowconfigure(1, weight=1)

        values_toolbar = ttk.Frame(values_frame)
        values_toolbar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        self.paste_btn = ttk.Button(
            values_toolbar,
            text="从剪贴板粘贴一列",
            command=self._paste_column_from_clipboard,
        )
        self.paste_btn.pack(side="left")
        ttk.Label(
            values_toolbar,
            text="可在 Excel 中复制整列后点击左侧按钮，也可直接在下方按 Ctrl+V",
            foreground="#666666",
        ).pack(side="left", padx=(8, 0))

        self.values_text = tk.Text(values_frame, height=7, wrap="none", undo=True)
        self.values_text.grid(row=1, column=0, sticky="nsew")
        values_y_scroll = ttk.Scrollbar(values_frame, orient="vertical", command=self.values_text.yview)
        values_y_scroll.grid(row=1, column=1, sticky="ns")
        values_x_scroll = ttk.Scrollbar(
            values_frame, orient="horizontal", command=self.values_text.xview
        )
        values_x_scroll.grid(row=2, column=0, sticky="ew")
        self.values_text.configure(
            yscrollcommand=values_y_scroll.set, xscrollcommand=values_x_scroll.set
        )
        self.values_text.bind("<Control-Return>", self._on_values_ctrl_enter)

        button_frame = ttk.Frame(main)
        button_frame.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        self.execute_btn = ttk.Button(button_frame, text="执行查询", command=self._execute_query)
        self.execute_btn.pack(side="left")
        self.clear_btn = ttk.Button(button_frame, text="清空", command=self._clear_all)
        self.clear_btn.pack(side="left", padx=(8, 0))
        self.export_btn = ttk.Button(button_frame, text="导出 Excel", command=self._export_excel)
        self.export_btn.pack(side="left", padx=(8, 0))
        ttk.Label(
            button_frame,
            text="每批最多 200 个值，分批查询后合并去重；Ctrl+Enter 可执行查询",
            foreground="#666666",
        ).pack(side="right")

        result_frame = ttk.LabelFrame(main, text="查询结果", padding=8)
        result_frame.grid(row=4, column=0, sticky="nsew", pady=(0, 8))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(result_frame, columns=(), show="headings", selectmode="extended")
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_y_scroll = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        tree_y_scroll.grid(row=0, column=1, sticky="ns")
        tree_x_scroll = ttk.Scrollbar(result_frame, orient="horizontal", command=self.tree.xview)
        tree_x_scroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=tree_y_scroll.set, xscrollcommand=tree_x_scroll.set)

        status_label = tk.Label(
            main,
            textvariable=self.status_var,
            anchor="w",
            relief="sunken",
            bd=1,
            padx=8,
            pady=4,
        )
        status_label.grid(row=5, column=0, sticky="ew")

    def _on_values_ctrl_enter(self, _event):
        self._execute_query()
        return "break"

    def _ui(self, callback, *args):
        def invoke():
            try:
                callback(*args)
            except tk.TclError:
                pass

        try:
            self.root.after(0, invoke)
        except tk.TclError:
            pass

    def _set_status(self, text):
        self.status_var.set(text)

    def _set_busy(self, busy):
        self._busy = busy
        state = "disabled" if busy else "normal"
        for button in (
            self.test_btn,
            self.save_conn_btn,
            self.clear_saved_btn,
            self.execute_btn,
            self.clear_btn,
        ):
            button.configure(state=state)
        for entry in (self.user_entry, self.password_entry, self.dsn_entry):
            entry.configure(state=state)
        self.remember_check.configure(state=state)
        if busy or not self._result_ready:
            self.export_btn.configure(state="disabled")
        else:
            self.export_btn.configure(state="normal")

    def _load_saved_connection(self):
        config = load_connection_config()
        if not config:
            return
        user = str(config.get("user") or "")
        dsn = str(config.get("dsn") or "")
        remember = bool(config.get("remember"))
        if user:
            self.user_var.set(user)
        if dsn:
            self.dsn_var.set(dsn)
        self.remember_var.set(remember)
        if remember:
            encrypted = str(config.get("password_enc") or "")
            if encrypted:
                try:
                    self.password_var.set(unprotect_password(encrypted))
                    self._set_status("已加载保存的连接信息和密码")
                except Exception:
                    self._set_status("已加载连接信息，但保存的密码无法解密，请重新输入")
            else:
                self._set_status("已加载保存的连接信息")
        else:
            self._set_status("已加载保存的用户名和连接串")

    def _save_connection(self):
        user = self.user_var.get().strip()
        dsn = self.dsn_var.get().strip()
        password = self.password_var.get()
        remember = self.remember_var.get()
        if not user:
            messagebox.showwarning("提示", "请先填写数据库用户名。", parent=self.root)
            return
        if not dsn:
            messagebox.showwarning("提示", "请先填写连接串。", parent=self.root)
            return
        if remember and not password:
            messagebox.showwarning(
                "提示",
                "已勾选“记住密码”，但密码为空，请填写密码后再保存。",
                parent=self.root,
            )
            return
        try:
            save_connection_config(user, dsn, password, remember)
        except Exception as exc:
            messagebox.showerror(
                "保存失败", f"保存连接信息失败：\n{exc}", parent=self.root
            )
            return
        if remember:
            self._set_status(f"连接信息已保存（密码已加密）：{CONFIG_PATH}")
        else:
            self._set_status(f"连接信息已保存（未保存密码）：{CONFIG_PATH}")

    def _clear_saved_connection(self):
        if not messagebox.askyesno(
            "确认", "确定要清除本机保存的连接信息吗？", parent=self.root
        ):
            return
        try:
            clear_connection_config()
        except Exception as exc:
            messagebox.showerror(
                "清除失败", f"清除保存的连接信息失败：\n{exc}", parent=self.root
            )
            return
        self.remember_var.set(False)
        self.password_var.set("")
        self._set_status("已清除本机保存的连接信息")

    def _auto_save_if_remembered(self):
        if not self.remember_var.get():
            return True
        try:
            save_connection_config(
                self.user_var.get().strip(),
                self.dsn_var.get().strip(),
                self.password_var.get(),
                True,
            )
            return True
        except Exception:
            return False

    def _paste_column_from_clipboard(self):
        try:
            content = self.root.clipboard_get()
        except tk.TclError:
            messagebox.showwarning(
                "剪贴板为空",
                "没有读取到文本内容。请先在 Excel 中复制一列，再点击此按钮。",
                parent=self.root,
            )
            return
        values, multiple_columns = parse_clipboard_column(content)
        if not values:
            messagebox.showwarning(
                "没有可粘贴的值",
                "剪贴板中没有可用的文本内容，请重新复制 Excel 列。",
                parent=self.root,
            )
            return
        unique_values = list(dict.fromkeys(values))
        self.values_text.delete("1.0", "end")
        self.values_text.insert("1.0", "\n".join(unique_values))
        status = f"已从剪贴板粘贴 {len(unique_values)} 个值"
        if multiple_columns:
            status += "（检测到多列，已取第一列）"
        if len(unique_values) < len(values):
            status += f"，忽略 {len(values) - len(unique_values)} 个重复值"
        self._set_status(status)

    def _collect_conn_info(self):
        user = self.user_var.get().strip()
        password = self.password_var.get()
        dsn = self.dsn_var.get().strip()
        if not user:
            raise ValueError("请填写数据库用户名。")
        if not dsn:
            raise ValueError("请填写连接串，例如 host:1521/service_name 或 TNS 别名。")
        return user, password, dsn

    def _require_identifier(self, value, label):
        if not value:
            raise ValueError(f"请填写{label}。")
        if not IDENT_RE.fullmatch(value):
            raise ValueError(f"{label}“{value}”不合法，仅允许字母、数字、下划线和点号。")
        return value

    def _parse_select_fields(self, raw):
        raw = raw.strip()
        if not raw:
            return []
        fields = []
        seen = set()
        for part in re.split(r"[,，]", raw):
            part = part.strip()
            if not part:
                continue
            if not IDENT_RE.fullmatch(part):
                raise ValueError(
                    f"查询字段“{part}”不合法，仅允许字母、数字、下划线和点号。"
                )
            key = part.lower()
            if key in seen:
                continue
            seen.add(key)
            fields.append(part)
        if not fields:
            raise ValueError("查询字段格式不正确，请用逗号分隔字段名，或留空使用 SELECT *。")
        return fields

    def _parse_values(self):
        raw = self.values_text.get("1.0", "end-1c")
        values = []
        seen = set()
        for line in raw.splitlines():
            value = line.strip()
            if not value or value in seen:
                continue
            seen.add(value)
            values.append(value)
        if not values:
            raise ValueError("请粘贴至少一个待匹配值，每行一个。")
        return values

    def _test_connection(self):
        try:
            conn_info = self._collect_conn_info()
        except ValueError as exc:
            messagebox.showwarning("输入有误", str(exc), parent=self.root)
            return
        self._set_busy(True)
        self._set_status("正在测试数据库连接...")
        threading.Thread(
            target=self._test_worker, args=conn_info, daemon=True
        ).start()

    def _test_worker(self, user, password, dsn):
        connection = None
        try:
            connection = oracledb.connect(user=user, password=password, dsn=dsn)
            cursor = connection.cursor()
            cursor.execute("SELECT 1 FROM DUAL")
            cursor.fetchone()
            cursor.close()
            self._ui(self._on_test_done)
        except Exception as exc:
            self._ui(self._on_worker_error, "连接失败", friendly_db_error(exc))
        finally:
            if connection is not None:
                try:
                    connection.close()
                except Exception:
                    pass

    def _on_test_done(self):
        self._set_busy(False)
        auto_saved = self._auto_save_if_remembered()
        if self.remember_var.get() and not auto_saved:
            self._set_status("连接测试成功，但连接信息自动保存失败")
        else:
            self._set_status("连接测试成功")
        messagebox.showinfo("连接测试", "数据库连接成功。", parent=self.root)

    def _execute_query(self):
        if self._busy:
            return
        try:
            conn_info = self._collect_conn_info()
            table = self._require_identifier(self.table_var.get().strip(), "表名")
            match_field = self._require_identifier(self.match_var.get().strip(), "匹配字段")
            select_fields = self._parse_select_fields(self.select_fields_var.get())
            values = self._parse_values()
        except ValueError as exc:
            messagebox.showwarning("输入有误", str(exc), parent=self.root)
            return

        self._reset_results()
        self._set_busy(True)
        batch_count = (len(values) + BATCH_SIZE - 1) // BATCH_SIZE
        self._set_status(
            f"准备查询：共 {len(values)} 个值，分 {batch_count} 批执行..."
        )
        threading.Thread(
            target=self._query_worker,
            args=(conn_info, table, match_field, select_fields, values),
            daemon=True,
        ).start()

    def _query_worker(self, conn_info, table, match_field, select_fields, values):
        connection = None
        try:
            self._ui(self._set_status, "正在连接数据库...")
            connection = oracledb.connect(
                user=conn_info[0], password=conn_info[1], dsn=conn_info[2]
            )
            cursor = connection.cursor()
            cursor.arraysize = 500

            if select_fields:
                select_clause = f"{', '.join(select_fields)}, {match_field}"
            else:
                select_clause = "*"

            batches = [
                values[start : start + BATCH_SIZE]
                for start in range(0, len(values), BATCH_SIZE)
            ]
            total_batches = len(batches)
            merged_rows = {}
            hit_map = {}
            matched_values = set()
            output_columns = None
            hit_column = "命中值"

            for batch_index, batch in enumerate(batches, start=1):
                self._ui(
                    self._set_status,
                    f"正在查询第 {batch_index}/{total_batches} 批（本批 {len(batch)} 个值）...",
                )
                conditions = " OR ".join(
                    f"{match_field} LIKE :{position}"
                    for position in range(1, len(batch) + 1)
                )
                sql = f"SELECT {select_clause} FROM {table} WHERE ({conditions})"
                binds = [f"%{value}%" for value in batch]
                cursor.execute(sql, binds)
                raw_columns = [description[0] for description in cursor.description]
                fetched_rows = cursor.fetchall()

                if select_fields:
                    match_index = len(raw_columns) - 1
                    display_indexes = list(range(len(raw_columns) - 1))
                else:
                    match_index = next(
                        (
                            index
                            for index, name in enumerate(raw_columns)
                            if column_key(name) == column_key(match_field)
                        ),
                        None,
                    )
                    if match_index is None:
                        raise ValueError(
                            f"匹配字段“{match_field}”不在 SELECT * 的结果中，请检查字段名。"
                        )
                    display_indexes = list(range(len(raw_columns)))

                if output_columns is None:
                    output_columns = [raw_columns[index] for index in display_indexes]
                    while hit_column in output_columns:
                        hit_column += "_"

                patterns = [(value, like_regex(value)) for value in batch]
                for raw_row in fetched_rows:
                    row_values = [normalize_value(value) for value in raw_row]
                    match_value = row_values[match_index]
                    match_text = "" if match_value is None else str(match_value)
                    hits = []
                    for value, pattern in patterns:
                        if pattern.match(match_text):
                            hits.append(value)
                            matched_values.add(value)

                    freezed_row = tuple(freeze_value(value) for value in row_values)
                    if freezed_row not in merged_rows:
                        merged_rows[freezed_row] = [
                            row_values[index] for index in display_indexes
                        ]
                        hit_map[freezed_row] = []
                    if hits:
                        bucket = hit_map[freezed_row]
                        for hit in hits:
                            if hit not in bucket:
                                bucket.append(hit)

            self._ui(self._set_status, "查询完成，正在整理结果...")
            output_rows = []
            for key, row_values in merged_rows.items():
                hit_text = " | ".join(hit_map[key]) if hit_map[key] else "（批内匹配）"
                output_rows.append(row_values + [hit_text])
            unmatched_values = [value for value in values if value not in matched_values]
            final_columns = output_columns + [hit_column]
            self._ui(self._on_query_done, final_columns, output_rows, unmatched_values)
        except Exception as exc:
            self._ui(self._on_worker_error, "查询失败", friendly_db_error(exc))
        finally:
            if connection is not None:
                try:
                    connection.close()
                except Exception:
                    pass

    def _on_query_done(self, columns, rows, unmatched_values):
        self._result_columns = columns
        self._result_rows = rows
        self._unmatched_values = unmatched_values
        self._result_ready = True
        self._set_busy(False)
        auto_saved = self._auto_save_if_remembered()

        if rows:
            self._status_summary = (
                f"查询完成：共 {len(rows)} 行，未匹配值 {len(unmatched_values)} 个"
            )
        else:
            self._status_summary = (
                f"查询完成：没有匹配到任何结果，未匹配值 {len(unmatched_values)} 个"
            )
        if len(rows) > MAX_DISPLAY_ROWS:
            self._status_summary += (
                f"；界面仅显示前 {MAX_DISPLAY_ROWS} 行，导出包含全部结果"
            )
        if self.remember_var.get() and not auto_saved:
            self._status_summary += "；连接信息自动保存失败"

        self._setup_tree(columns)
        self._render_token += 1
        token = self._render_token
        self._set_status(self._status_summary)
        display_rows = rows[:MAX_DISPLAY_ROWS]
        if display_rows:
            self._set_status("正在加载结果到界面...")
            self._render_chunk(display_rows, 0, token)
        else:
            self._set_status(self._status_summary)

    def _setup_tree(self, columns):
        for item in self.tree.get_children(""):
            self.tree.delete(item)
        if not columns:
            self.tree.configure(columns=())
            return
        column_ids = [f"column_{index}" for index in range(len(columns))]
        self.tree.configure(columns=column_ids)
        for column_id, title in zip(column_ids, columns):
            width = 220 if str(title).startswith("命中值") else 150
            self.tree.heading(column_id, text=str(title), anchor="w")
            self.tree.column(
                column_id, width=width, minwidth=60, anchor="w", stretch=False
            )

    def _render_chunk(self, rows, start, token):
        if token != self._render_token:
            return
        end = min(start + 150, len(rows))
        for row in rows[start:end]:
            self.tree.insert("", "end", values=tuple(cell_text(value) for value in row))
        if end < len(rows):
            self.root.after(1, lambda: self._render_chunk(rows, end, token))
        else:
            self._set_status(self._status_summary)

    def _on_worker_error(self, title, detail):
        self._set_busy(False)
        self._set_status(f"{title}，请查看错误提示")
        messagebox.showerror(title, detail, parent=self.root)

    def _reset_results(self):
        self._render_token += 1
        self._result_ready = False
        self._result_columns = []
        self._result_rows = []
        self._unmatched_values = []
        self._status_summary = ""
        self._setup_tree([])
        self.export_btn.configure(state="disabled")

    def _clear_all(self):
        if self._busy:
            return
        self.values_text.delete("1.0", "end")
        self._reset_results()
        self._set_status("已清空待匹配值和查询结果")

    def _export_excel(self):
        if self._busy or not self._result_ready:
            messagebox.showwarning("提示", "请先执行查询，再导出结果。", parent=self.root)
            return
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="导出 Excel",
            defaultextension=".xlsx",
            initialfile=f"查询结果_{timestamp}.xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        columns = list(self._result_columns)
        rows = [list(row) for row in self._result_rows]
        unmatched_values = list(self._unmatched_values)
        self._set_busy(True)
        self._set_status("正在导出 Excel...")
        threading.Thread(
            target=self._export_worker,
            args=(path, columns, rows, unmatched_values),
            daemon=True,
        ).start()

    def _export_worker(self, path, columns, rows, unmatched_values):
        try:
            safe_rows = [
                [excel_safe(value) for value in row]
                for row in rows
            ]
            result_df = pd.DataFrame(safe_rows, columns=columns)
            unmatched_df = pd.DataFrame({"未匹配值": unmatched_values})
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                result_df.to_excel(writer, sheet_name="查询结果", index=False)
                unmatched_df.to_excel(writer, sheet_name="未匹配值", index=False)
                for worksheet in writer.book.worksheets:
                    autofit_sheet(worksheet)
            self._ui(self._on_export_done, path)
        except Exception as exc:
            self._ui(self._on_worker_error, "导出失败", f"导出 Excel 失败：\n{exc}")

    def _on_export_done(self, path):
        self._set_busy(False)
        self._set_status(f"导出成功：{path}")
        messagebox.showinfo("导出成功", f"Excel 已导出到：\n{path}", parent=self.root)


def main():
    root = tk.Tk()
    OracleBatchQueryApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

# 启动方式：
#   pip install -r requirements.txt
#   python oracle_batch_query.py
#
# 打包 exe：
#   双击 build_exe.bat，或运行 build_exe.ps1，生成 dist\OracleBatchQuery.exe
#
# 易错配置说明：
# 1. 连接串常用格式为 主机:端口/服务名（如 127.0.0.1:1521/orclpdb1），也可填写本机 tnsnames.ora 中配置的 TNS 别名。
# 2. 默认使用 python-oracledb 的 Thin 模式，一般不需要 Oracle Instant Client；若必须使用 Thick 模式，
#    请在 import oracledb 之后、建立连接之前调用 oracledb.init_oracle_client(lib_dir=r"客户端路径")，
#    并确保 Instant Client 位数与 Python 一致。
# 3. Oracle 中未加双引号创建的表名、字段名默认按大写存储；这里填写小写通常也能查询，但若建表时用了双引号，请按实际大小写填写。
# 4. 表名、匹配字段、查询字段仅允许字母、数字、下划线和点号，不支持别名、函数、带引号的标识符或 SQL 片段。
# 5. 粘贴值中的 % 和 _ 会被 Oracle 当作 LIKE 通配符；本工具按需求不做转义，搜索普通 %、_ 字符时请注意。
# 6. 每批最多 200 个值；结果较多时界面只展示前 2000 行，导出的 Excel 包含全部结果。
# 7. 密码中的特殊字符无需转义；请勿把真实连接信息、密码写进代码或提交到版本库。
# 8. Excel 复制一列后，可直接在待匹配值区域 Ctrl+V，或点击“从剪贴板粘贴一列”；检测到多列时只取第一列。
# 9. 勾选“记住密码”后点击“保存连接”即可保存；测试连接或查询成功后也会自动保存。密码使用 Windows DPAPI
#    按当前 Windows 用户加密，配置文件位于 %APPDATA%\OracleBatchQuery\connection.json，换用户或换电脑后无法解密。
# 10. 打包后的 exe 首次启动会解压运行文件，可能比直接运行 .py 慢几秒，属于 PyInstaller 单文件模式的正常现象。
