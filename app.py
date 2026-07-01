"""MailBatch – pywebview single-window application."""

from __future__ import annotations

import json
import sys
import threading
import webbrowser
from datetime import datetime
from pathlib import Path

import webview

from app_config import (
    AppConfig,
    config_path,
    export_app_config,
    import_app_config,
    load_app_config,
    save_app_config,
)
from smtp_password_help import get_password_help
from smtp_presets import (
    find_preset_by_domain,
    get_preset,
    preset_combo_items,
)
from mail_core import (
    SmtpConfig,
    build_email_preview,
    export_failed_send_results,
    find_cc_column,
    find_email_column,
    load_excel,
    send_batch,
    validate_send_request,
)


def _app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _app_html_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / "app.html"  # type: ignore[attr-defined]
    return _app_dir() / "app.html"


APP_DIR = _app_dir()
LAST_CONFIG_FILE = config_path(APP_DIR)


def _prompt_save_on_exit() -> str:
    """Ask whether to save config on exit. Returns 'save', 'discard', or 'cancel'."""
    message = (
        "是否保存当前配置？\n\n"
        "「是」保存并退出\n"
        "「否」不保存直接退出\n"
        "「取消」返回程序"
    )
    title = "退出 MailBatch"

    if sys.platform == "win32":
        import ctypes

        result = ctypes.windll.user32.MessageBoxW(  # type: ignore[attr-defined]
            0,
            message,
            title,
            0x00000003 | 0x00000040,  # MB_YESNOCANCEL | MB_ICONINFORMATION
        )
        if result == 6:  # IDYES
            return "save"
        if result == 7:  # IDNO
            return "discard"
        return "cancel"

    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    choice = messagebox.askyesnocancel(title, message, parent=root)
    root.destroy()
    if choice is True:
        return "save"
    if choice is False:
        return "discard"
    return "cancel"


def _show_error_message(text: str) -> None:
    if sys.platform == "win32":
        import ctypes

        ctypes.windll.user32.MessageBoxW(0, text, "MailBatch", 0x00000010)  # type: ignore[attr-defined]
        return

    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("MailBatch", text, parent=root)
    root.destroy()


class AppAPI:
    """Python back-end exposed to JavaScript via pywebview."""

    def __init__(self) -> None:
        self._excel_data = None
        self._excel_path: str = ""
        self._send_worker: threading.Thread | None = None
        self._stop_send = False
        self._config_snapshot: dict | None = None

    # ── Config ────────────────────────────────────────────────────────────

    def load_config(self) -> dict:
        if not LAST_CONFIG_FILE.is_file():
            return {"_has_saved_config": False}
        try:
            cfg = load_app_config(LAST_CONFIG_FILE).to_dict()
            self._config_snapshot = cfg
            cfg["_has_saved_config"] = True
            return cfg
        except Exception as exc:  # noqa: BLE001
            return {"error": str(exc), "_has_saved_config": False}

    def sync_config(self, config: dict) -> None:
        """Store latest UI state in memory (no disk I/O). Called from JS on change."""
        self._config_snapshot = config

    def save_config(self, config: dict) -> dict:
        try:
            self._config_snapshot = config
            save_app_config(AppConfig.from_dict(config), LAST_CONFIG_FILE)
            return {"ok": True}
        except Exception as exc:  # noqa: BLE001
            return {"error": str(exc)}

    def export_config(self, config: dict) -> dict:
        win = webview.windows[0]
        paths = win.create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename="mailbatch_config.json",
            file_types=("JSON files (*.json)",),
        )
        if not paths:
            return {"cancelled": True}
        try:
            export_app_config(AppConfig.from_dict(config), paths[0])
            return {"ok": True, "path": paths[0]}
        except Exception as exc:  # noqa: BLE001
            return {"error": str(exc)}

    def import_config(self) -> dict:
        win = webview.windows[0]
        paths = win.create_file_dialog(
            webview.OPEN_DIALOG,
            file_types=("JSON files (*.json)",),
        )
        if not paths:
            return {"cancelled": True}
        try:
            return import_app_config(paths[0]).to_dict()
        except Exception as exc:  # noqa: BLE001
            return {"error": str(exc)}

    def reset_session(self) -> None:
        """Clear in-memory session state without touching config files on disk."""
        self._excel_data = None
        self._excel_path = ""
        self._config_snapshot = None

    # ── SMTP presets ──────────────────────────────────────────────────────

    def get_smtp_presets(self) -> list:
        return [{"id": pid, "label": label} for pid, label in preset_combo_items()]

    def get_smtp_preset_details(self, preset_id: str) -> dict | None:
        preset = get_preset(preset_id)
        if preset is None:
            return None
        return {
            "host": preset.host,
            "port": preset.port,
            "use_tls": preset.use_tls,
            "use_ssl": preset.use_ssl,
        }

    def detect_smtp(self, email: str) -> dict | None:
        preset = find_preset_by_domain(email)
        if preset is None:
            return None
        return {
            "host": preset.host,
            "port": preset.port,
            "use_tls": preset.use_tls,
            "use_ssl": preset.use_ssl,
            "label": preset.label,
        }

    def get_password_help_url(self, email: str, preset_id: str) -> dict:
        info = get_password_help(email, preset_id)
        return {"url": info.help_url, "provider": info.provider}

    def open_url(self, url: str) -> None:
        webbrowser.open(url)

    # ── Excel ─────────────────────────────────────────────────────────────

    def choose_excel(self) -> dict:
        win = webview.windows[0]
        paths = win.create_file_dialog(
            webview.OPEN_DIALOG,
            file_types=("Excel files (*.xlsx;*.xlsm)", "All files (*.*)"),
        )
        if not paths:
            return {}
        return self._load_excel(paths[0])

    def reload_excel(self, path: str) -> dict:
        if not Path(path).is_file():
            return {"error": f"文件不存在：{path}"}
        return self._load_excel(path)

    def _load_excel(self, path: str) -> dict:
        try:
            data = load_excel(path)
            self._excel_data = data
            self._excel_path = path
            headers = [h for h in data.headers if h]
            return {
                "path": path,
                "sheet": data.sheet_name,
                "rows": len(data.rows),
                "headers": headers,
                "email_col": find_email_column(headers) or "",
                "cc_col": find_cc_column(headers) or "",
            }
        except Exception as exc:  # noqa: BLE001
            return {"error": str(exc)}

    def _ensure_excel_loaded(self, excel_path: str = "") -> dict | None:
        if self._excel_data is not None:
            return None
        path = excel_path or self._excel_path
        if not path:
            return {"error": "未加载 Excel 数据，请先选择 Excel 文件。"}
        result = self._load_excel(path)
        if result.get("error"):
            return {"error": result["error"]}
        return None

    # ── Attachments ───────────────────────────────────────────────────────

    def add_attachments(self) -> list:
        win = webview.windows[0]
        paths = win.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=True,
            file_types=("All files (*.*)",),
        )
        return list(paths) if paths else []

    # ── Template ──────────────────────────────────────────────────────────

    def load_template_file(self) -> dict:
        win = webview.windows[0]
        paths = win.create_file_dialog(
            webview.OPEN_DIALOG,
            file_types=("Template files (*.html;*.txt)", "All files (*.*)"),
        )
        if not paths:
            return {}
        path = Path(paths[0])
        try:
            text = path.read_text(encoding="utf-8")
            if path.suffix.lower() != ".html":
                import html as _h
                text = (
                    "<pre style='font-family:inherit;white-space:pre-wrap'>"
                    + _h.escape(text)
                    + "</pre>"
                )
            return {"path": str(path), "html": text}
        except Exception as exc:  # noqa: BLE001
            return {"error": str(exc)}

    def save_template_file(self, html_content: str, current_path: str = "") -> str | None:
        win = webview.windows[0]
        cur = Path(current_path) if current_path else None
        save_name = (cur.name if cur and cur.suffix.lower() == ".html" else "template.html")
        paths = win.create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename=save_name,
            file_types=("HTML files (*.html)",),
        )
        if not paths:
            return None
        Path(paths[0]).write_text(html_content, encoding="utf-8")
        return paths[0]

    # ── Preview ───────────────────────────────────────────────────────────

    def preview_email(self, data: dict) -> dict:
        load_err = self._ensure_excel_loaded(data.get("excel_path", ""))
        if load_err:
            return load_err
        assert self._excel_data is not None
        row_index = int(data.get("row_index") or 1)
        try:
            preview = build_email_preview(
                self._excel_data,
                row_index,
                data["template"],
                data["subject"],
                data.get("attachment_paths", []),
                email_column=data.get("email_column") or None,
                cc_column=data.get("cc_column") or None,
            )
        except Exception as exc:  # noqa: BLE001
            return {"error": str(exc)}

        return {
            "ok": True,
            "row_index": row_index,
            "total_rows": len(self._excel_data.rows),
            "to_email": preview.to_email,
            "cc_emails": preview.cc_emails,
            "subject_template": data.get("subject", ""),
            "subject": preview.subject,
            "body": preview.body,
            "attachments": [Path(p).name for p in preview.attachment_paths],
        }

    # ── Validate ──────────────────────────────────────────────────────────

    def validate_send(self, data: dict) -> dict:
        load_err = self._ensure_excel_loaded(data.get("excel_path", ""))
        if load_err:
            return {"errors": [load_err["error"]], "warnings": []}
        assert self._excel_data is not None
        issues = validate_send_request(
            self._excel_data,
            data["template"],
            data["subject"],
            data.get("attachment_paths", []),
            email_column=data.get("email_column") or None,
            cc_column=data.get("cc_column") or None,
        )
        return {
            "errors": [i.message for i in issues if i.level == "error"],
            "warnings": [i.message for i in issues if i.level == "warning"],
        }

    # ── Send ──────────────────────────────────────────────────────────────

    def start_send(self, data: dict) -> dict:
        if self._send_worker is not None and self._send_worker.is_alive():
            return {"error": "发送中，请等待"}
        load_err = self._ensure_excel_loaded(data.get("excel_path", ""))
        if load_err:
            return load_err
        assert self._excel_data is not None

        smtp_config = SmtpConfig(
            host=data["smtp_host"],
            port=int(data.get("smtp_port", 587)),
            user=data.get("smtp_user", "") or data.get("smtp_email", ""),
            password=data.get("smtp_password", ""),
            from_email=data.get("smtp_email", "") or data.get("smtp_user", ""),
            use_tls=bool(data.get("use_tls", True)),
            use_ssl=bool(data.get("use_ssl", False)),
            send_interval=float(data.get("send_interval", 1.0)),
        )

        excel_data = self._excel_data
        template = data["template"]
        subject = data["subject"]
        attachments = data.get("attachment_paths", [])
        email_column = data.get("email_column") or None
        cc_column = data.get("cc_column") or None

        self._stop_send = False

        def _progress(current: int, total: int, result) -> None:
            try:
                r = {
                    "row_index": result.row_index,
                    "recipient": result.recipient,
                    "success": result.success,
                    "error": result.error,
                }
                webview.windows[0].evaluate_js(
                    f"onProgress({current},{total},{json.dumps(r)})"
                )
            except Exception:  # noqa: BLE001
                pass

        def _log(text: str) -> None:
            try:
                webview.windows[0].evaluate_js(f"appendLog({json.dumps(text)})")
            except Exception:  # noqa: BLE001
                pass

        def _run() -> None:
            try:
                summary = send_batch(
                    excel_data, template, subject, smtp_config, attachments,
                    email_column=email_column,
                    cc_column=cc_column,
                    progress_callback=_progress,
                    log_callback=_log,
                    should_stop=lambda: self._stop_send,
                )
                failed_path: str | None = None
                if (summary.failure_count + summary.skipped_count) > 0:
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    out = APP_DIR / f"发送失败_{ts}.xlsx"
                    try:
                        failed_path = str(export_failed_send_results(excel_data, summary, out))
                    except Exception:  # noqa: BLE001
                        pass

                s = {
                    "success_count": summary.success_count,
                    "failure_count": summary.failure_count,
                    "skipped_count": summary.skipped_count,
                    "total": summary.total,
                    "failed_path": failed_path,
                }
                webview.windows[0].evaluate_js(f"onSendFinished({json.dumps(s)})")
            except Exception as exc:  # noqa: BLE001
                webview.windows[0].evaluate_js(f"onSendFailed({json.dumps(str(exc))})")
            finally:
                self._send_worker = None

        self._send_worker = threading.Thread(target=_run, daemon=True)
        self._send_worker.start()
        return {"started": True}

    def cancel_send(self) -> None:
        self._stop_send = True


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────

def main() -> int:
    html_path = _app_html_path()
    if not html_path.exists():
        print(f"ERROR: app.html not found at {html_path}", file=sys.stderr)
        return 1

    html_content = html_path.read_text(encoding="utf-8")
    api = AppAPI()

    win = webview.create_window(
        "MailBatch 批量邮件工具",
        html=html_content,
        js_api=api,
        width=1300,
        height=820,
        min_size=(960, 640),
    )

    def _on_closing() -> bool:
        api._stop_send = True
        choice = _prompt_save_on_exit()
        if choice == "cancel":
            return False
        if choice == "save":
            if not api._config_snapshot:
                return True
            try:
                snapshot = {
                    key: value
                    for key, value in api._config_snapshot.items()
                    if not str(key).startswith("_")
                }
                save_app_config(AppConfig.from_dict(snapshot), LAST_CONFIG_FILE)
            except Exception as exc:  # noqa: BLE001
                _show_error_message(f"配置保存失败：{exc}")
                return False
        return True

    win.events.closing += _on_closing

    icon_path = APP_DIR / "assets" / "mailbatch.ico"
    if getattr(sys, "frozen", False):
        meipass_icon = Path(sys._MEIPASS) / "assets" / "mailbatch.ico"  # type: ignore[attr-defined]
        if meipass_icon.exists():
            icon_path = meipass_icon

    webview.start(icon=str(icon_path) if icon_path.exists() else None)
    return 0


if __name__ == "__main__":
    sys.exit(main())
