"""Invoice Renamer — WebView 版主界面入口

保留原有业务逻辑，仅将 UI 层替换为 PyWebview + HTML/CSS/JS。
运行时当前目录需在项目根目录。
"""

import json
import os
import sys
import threading
import traceback
from pathlib import Path

# 项目根目录
ROOT = Path(__file__).parent.resolve()
WEBVIEW_DIR = ROOT / "webview"

# 复用现有业务模块
sys.path.insert(0, str(ROOT))
from config import (
    APP_AUTHOR, APP_ID, APP_VERSION,
    DEFAULT_FIELD_ENABLED, DEFAULT_FIELD_ORDER,
    FIELD_LABELS, RESULT_HINT,
    ALLOWED_EXTENSIONS,
    STAT_CARD_STYLE, TREE_TAG_COLORS,
)
from invoice_parser import parse_invoice, parse_image_cloud
from name_builder import build_name, ensure_unique_name
from excel_exporter import export_invoice_excel


def resource_path(relative_path: str) -> str:
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)


class Api:
    """暴露给前端 JS 调用的 Python API。"""

    def __init__(self):
        self.window = None
        self.selected_paths: list[str] = []
        self.records: list[dict] = []
        self.processing: bool = False
        self._scanning: bool = False
        self.rename_history: list[tuple[str, str, str]] = []
        self.field_order: list[str] = list(DEFAULT_FIELD_ORDER)
        self.field_enabled: dict[str, bool] = dict(DEFAULT_FIELD_ENABLED)
        self.custom_value: str = ""
        self.preview_mode: bool = True
        self._cloud_ocr_enabled: bool = False
        self._cloud_secret_id: str = ""
        self._cloud_secret_key: str = ""
        self._load_cloud_ocr_state()

    def open_browser(self, url: str) -> dict:
        """用系统默认浏览器打开链接。"""
        try:
            import os
            os.startfile(url)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ── 配置与状态 ──────────────────────────────────────────────────────

    def ping(self, message: str = "ping") -> dict:
        """前端通信测试。"""
        try:
            with open("webview_debug.log", "a", encoding="utf-8") as f:
                f.write(f"[python] ping received: {message}\n")
        except Exception:
            pass
        return {"ok": True, "echo": message}


    def _load_cloud_ocr_state(self) -> None:
        try:
            from cloud_ocr import load_credentials
            creds = load_credentials()
            self._cloud_ocr_enabled = creds.get("enabled", False)
            self._cloud_secret_id = creds.get("secret_id", "")
            self._cloud_secret_key = creds.get("secret_key", "")
        except Exception:
            pass

    def _has_cloud_creds(self) -> bool:
        return bool(self._cloud_secret_id and self._cloud_secret_key)

    def get_init_state(self) -> dict:
        """前端初始化时获取完整状态。"""
        return {
            "version": APP_VERSION,
            "author": APP_AUTHOR,
            "result_hint": RESULT_HINT,
            "field_labels": FIELD_LABELS,
            "field_order": self.field_order,
            "field_enabled": self.field_enabled,
            "custom_value": self.custom_value,
            "preview_mode": self.preview_mode,
            "cloud": {
                "enabled": self._cloud_ocr_enabled and self._has_cloud_creds(),
                "configured": self._has_cloud_creds(),
                "secret_id": self._cloud_secret_id,
            },
            "stats": self._calc_stats(),
            "records": self._serialize_records(),
        }

    def _calc_stats(self) -> dict:
        total = len(self.records)
        complete = sum(1 for r in self.records if r.get("status") == "complete")
        partial = sum(1 for r in self.records if r.get("status") == "partial")
        failed = sum(1 for r in self.records if r.get("status") == "failed")
        not_invoice = sum(1 for r in self.records if r.get("status") == "not_invoice")
        return {
            "total": total,
            "complete": complete,
            "partial": partial,
            "failed": failed,
            "not_invoice": not_invoice,
        }

    def _serialize_records(self) -> list[dict]:
        out = []
        for idx, r in enumerate(self.records, start=1):
            out.append({
                "idx": idx,
                "path": r.get("path", ""),
                "source_name": r.get("source_name", ""),
                "current_name": r.get("current_name", ""),
                "new_name": r.get("new_name", ""),
                "type": r.get("fields", {}).get("type", ""),
                "seller": r.get("fields", {}).get("seller", ""),
                "amount": r.get("fields", {}).get("amount", ""),
                "status": r.get("status", "idle"),
                "error": r.get("error", ""),
            })
        return out

    # ── 文件选择 ────────────────────────────────────────────────────────

    def choose_folder(self) -> dict:
        try:
            import webview
            if not self.window and webview.windows:
                self.window = webview.windows[0]
            if not self.window:
                return {"ok": False, "error": "窗口未就绪"}
            folder = self.window.create_file_dialog(webview.FOLDER_DIALOG, directory="")
            if folder and isinstance(folder, (list, tuple)):
                folder = folder[0]
            if folder:
                self.selected_paths = []
                for root_dir, _, files in os.walk(folder):
                    for name in files:
                        path = os.path.join(root_dir, name)
                        if path.lower().endswith(ALLOWED_EXTENSIONS):
                            self.selected_paths.append(path)
                return {"ok": True, "path": folder, "summary": f"{len(self.selected_paths)} 个文件"}
            return {"ok": False}
        except Exception as e:
            traceback.print_exc()
            return {"ok": False, "error": str(e)}

    def choose_files(self) -> dict:
        try:
            import webview
            if not self.window and webview.windows:
                self.window = webview.windows[0]
            if not self.window:
                return {"ok": False, "error": "窗口未就绪"}
            files = self.window.create_file_dialog(
                webview.OPEN_DIALOG,
                directory="",
                allow_multiple=True,
                file_types=(
                    "发票文件 (*.pdf;*.jpg;*.jpeg;*.png;*.bmp;*.tiff)",
                    "PDF 文件 (*.pdf)",
                    "图片文件 (*.jpg;*.jpeg;*.png;*.bmp;*.tiff)",
                ),
            )
            if files:
                self.selected_paths = list(files)
                ext_counts = {}
                for f in files:
                    ext = os.path.splitext(f)[1].lower()
                    ext_counts[ext] = ext_counts.get(ext, 0) + 1
                summary = ", ".join(f"{c} {e}" for e, c in ext_counts.items())
                return {"ok": True, "path": f"已选择 {len(files)} 个文件 ({summary})", "summary": summary}
            return {"ok": False}
        except Exception as e:
            traceback.print_exc()
            return {"ok": False, "error": str(e)}

    def clear_source(self) -> dict:
        self.selected_paths = []
        self.records = []
        self.rename_history.clear()
        self._scanning = False
        self.processing = False
        return {"ok": True, "state": self.get_init_state()}

    # ── 扫描识别 ────────────────────────────────────────────────────────

    def _resolve_input_paths(self) -> list[str]:
        if self.selected_paths:
            return [p for p in self.selected_paths if os.path.isfile(p) and p.lower().endswith(ALLOWED_EXTENSIONS)]
        return []

    def scan_files(self) -> None:
        """在后台线程扫描文件并向前端推送进度。"""
        if self._scanning or self.processing:
            return
        paths = self._resolve_input_paths()
        if not paths:
            self._emit("status", {"message": "请先选择文件或文件夹"})
            return

        self._scanning = True
        self._emit("scan_started", {"total": len(paths)})

        def worker():
            records = []
            for i, path in enumerate(paths):
                try:
                    ext = os.path.splitext(path)[1].lower()
                    if ext == ".pdf":
                        parsed = parse_invoice(path)
                    elif ext in (".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif"):
                        if self._cloud_ocr_enabled and self._has_cloud_creds():
                            parsed = parse_image_cloud(
                                path,
                                self._cloud_secret_id,
                                self._cloud_secret_key,
                            )
                        else:
                            parsed = {"fields": {}, "error": "图片识别需要启用云端 OCR", "not_invoice": True}
                    else:
                        parsed = {"fields": {}, "error": "不支持的文件类型", "not_invoice": True}

                    fields = parsed.get("fields", {})
                    parse_error = parsed.get("error", "")
                    is_not_invoice = parsed.get("not_invoice", False)

                    current_name = os.path.basename(path)
                    record = {
                        "path": path,
                        "source_name": current_name,
                        "current_name": current_name,
                        "new_name": current_name,
                        "fields": fields,
                    }

                    if is_not_invoice:
                        record["status"] = "not_invoice"
                        record["error"] = parse_error or "非发票文件"
                    elif parse_error or not self._has_required_fields(fields):
                        record["status"] = "failed"
                        record["error"] = parse_error or "识别失败"
                    else:
                        record["status"] = "complete"
                        record["error"] = ""

                    # 预生成新文件名
                    record["new_name"] = self._build_target_name(record)
                    records.append(record)

                except Exception as e:
                    records.append({
                        "path": path,
                        "source_name": os.path.basename(path),
                        "current_name": os.path.basename(path),
                        "new_name": os.path.basename(path),
                        "fields": {},
                        "status": "failed",
                        "error": str(e),
                    })

                self._emit("scan_progress", {"current": i + 1, "total": len(paths)})

            self.records = records
            self._scanning = False
            self._emit("scan_finished", {
                "records": self._serialize_records(),
                "stats": self._calc_stats(),
                "message": f"识别完成，共 {len(records)} 个文件",
            })

        threading.Thread(target=worker, daemon=True).start()

    def _has_required_fields(self, fields: dict) -> bool:
        # 至少要有类型和金额，或者类型和日期
        return bool(fields.get("type")) and (bool(fields.get("amount")) or bool(fields.get("date")))

    def _build_target_name(self, record: dict) -> str:
        try:
            ext = os.path.splitext(record["path"])[1].lower() or ".pdf"
            return build_name(
                record["fields"],
                self.field_enabled,
                self.field_order,
                self.custom_value,
                ext=ext,
            )
        except Exception:
            return record["current_name"]

    # ── 模板与自定义字段 ────────────────────────────────────────────────

    def update_template(self, field_order: list[str], field_enabled: dict[str, bool], custom_value: str) -> dict:
        self.field_order = list(field_order)
        self.field_enabled = dict(field_enabled)
        self.custom_value = custom_value
        # 重新生成所有新文件名
        for r in self.records:
            if r.get("status") in ("complete", "partial"):
                r["new_name"] = self._build_target_name(r)
        return {"ok": True, "records": self._serialize_records()}

    def set_preview_mode(self, preview_mode: bool) -> dict:
        self.preview_mode = bool(preview_mode)
        return {"ok": True, "can_rename": not self.preview_mode and bool(self.records) and not self.processing}

    # ── 重命名 ──────────────────────────────────────────────────────────

    def start_rename(self) -> None:
        if self.processing or self._scanning or not self.records or self.preview_mode:
            return
        self.processing = True
        self._emit("rename_started", {"total": len(self.records)})

        records_snapshot = list(enumerate(self.records, start=1))
        total = len(records_snapshot)

        def worker():
            success = failed = skipped = 0
            errors = []
            history = []
            for idx, record in records_snapshot:
                if record.get("status") != "complete":
                    skipped += 1
                    self._emit("rename_progress", {
                        "current": idx, "total": total,
                        "success": success, "failed": failed, "skipped": skipped,
                    })
                    continue

                source_path = record["path"]
                directory = os.path.dirname(source_path)
                ext = os.path.splitext(source_path)[1].lower() or ".pdf"
                target_name = ensure_unique_name(
                    directory,
                    build_name(record["fields"], self.field_enabled, self.field_order, self.custom_value, ext=ext),
                    record["current_name"],
                )
                target_path = os.path.join(directory, target_name)
                try:
                    rename_happened = False
                    if record["current_name"] != target_name:
                        os.rename(source_path, target_path)
                        record["path"] = target_path
                        rename_happened = True
                    record["current_name"] = target_name
                    record["new_name"] = target_name
                    if rename_happened:
                        history.append((source_path, target_path, target_name))
                    success += 1
                    self._emit("rename_item_done", {
                        "idx": idx, "new_name": target_name, "status": "success"
                    })
                except Exception as e:
                    errors.append(f"文件: {record['source_name']}\n错误: {e}\n{traceback.format_exc()}\n")
                    failed += 1
                    self._emit("rename_item_done", {
                        "idx": idx, "new_name": record["current_name"], "status": "failed", "error": str(e)[:30]
                    })

                self._emit("rename_progress", {
                    "current": idx, "total": total,
                    "success": success, "failed": failed, "skipped": skipped,
                })

            self.rename_history.extend(history)
            self.processing = False
            if errors:
                with open("rename_errors.log", "w", encoding="utf-8") as f:
                    f.write("\n---\n".join(errors))
                msg = "重命名完成，部分失败详见 rename_errors.log"
            else:
                msg = "重命名完成"
            self._emit("rename_finished", {
                "records": self._serialize_records(),
                "stats": self._calc_stats(),
                "message": msg,
                "can_undo": bool(self.rename_history),
            })

        threading.Thread(target=worker, daemon=True).start()

    def undo_rename(self) -> None:
        if not self.rename_history or self.processing or self._scanning:
            return
        self.processing = True
        self._emit("undo_started", {})

        def worker():
            success = failed = 0
            for source_path, target_path, target_name in reversed(self.rename_history):
                if not os.path.exists(target_path):
                    failed += 1
                    continue
                try:
                    os.rename(target_path, source_path)
                    for record in self.records:
                        if record.get("path") == target_path:
                            record["path"] = source_path
                            record["current_name"] = os.path.basename(source_path)
                            record["new_name"] = os.path.basename(source_path)
                            break
                    success += 1
                except Exception:
                    failed += 1

            self.rename_history.clear()
            self.processing = False
            for r in self.records:
                if r.get("status") in ("complete",):
                    r["new_name"] = self._build_target_name(r)
            self._emit("undo_finished", {
                "records": self._serialize_records(),
                "stats": self._calc_stats(),
                "message": "撤销完成" if failed == 0 else "撤销完成，部分文件已不存在",
                "can_undo": False,
            })

        threading.Thread(target=worker, daemon=True).start()

    def on_rename_button_click(self) -> dict:
        if self.rename_history:
            self.undo_rename()
            return {"ok": True, "action": "undo"}
        else:
            self.start_rename()
            return {"ok": True, "action": "rename"}

    # ── Excel 导出 ──────────────────────────────────────────────────────

    def export_excel(self) -> dict:
        try:
            if not self.records:
                return {"ok": False, "error": "没有可导出的数据"}
            path = export_invoice_excel(self.records)
            return {"ok": True, "path": path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ── 云端 OCR 设置 ────────────────────────────────────────────────────

    def get_cloud_settings(self) -> dict:
        from cloud_ocr import get_usage_stats
        usage = get_usage_stats()
        return {
            "secret_id": self._cloud_secret_id,
            "secret_key": self._cloud_secret_key,
            "enabled": self._cloud_ocr_enabled and self._has_cloud_creds(),
            "usage": usage,
        }

    def save_cloud_settings(self, secret_id: str, secret_key: str, enabled: bool) -> dict:
        from cloud_ocr import save_credentials
        save_credentials(secret_id.strip(), secret_key.strip(), enabled=bool(enabled))
        self._cloud_secret_id = secret_id.strip()
        self._cloud_secret_key = secret_key.strip()
        self._cloud_ocr_enabled = bool(enabled)
        return {
            "ok": True,
            "cloud": {
                "enabled": self._cloud_ocr_enabled and self._has_cloud_creds(),
                "configured": self._has_cloud_creds(),
                "secret_id": self._cloud_secret_id,
            },
        }

    def clear_cloud_settings(self) -> dict:
        from cloud_ocr import clear_credentials
        clear_credentials()
        self._cloud_secret_id = ""
        self._cloud_secret_key = ""
        self._cloud_ocr_enabled = False
        return {
            "ok": True,
            "cloud": {
                "enabled": False,
                "configured": False,
                "secret_id": "",
            },
        }

    def verify_cloud_credentials(self, secret_id: str, secret_key: str) -> dict:
        from cloud_ocr import validate_credentials
        valid, msg = validate_credentials(secret_id.strip(), secret_key.strip())
        return {"ok": valid, "message": msg}

    def toggle_cloud_enabled(self) -> dict:
        if not self._has_cloud_creds():
            return {"ok": False, "error": "未配置密钥"}
        self._cloud_ocr_enabled = not self._cloud_ocr_enabled
        from cloud_ocr import save_credentials
        save_credentials(self._cloud_secret_id, self._cloud_secret_key, enabled=self._cloud_ocr_enabled)
        return {
            "ok": True,
            "cloud": {
                "enabled": self._cloud_ocr_enabled,
                "configured": True,
                "secret_id": self._cloud_secret_id,
            },
        }

    # ── 事件推送 ────────────────────────────────────────────────────────

    def _emit(self, event: str, data: dict) -> None:
        """通过 webview 的 js 桥接向前端发送事件。"""
        try:
            import webview
            if webview.windows:
                win = webview.windows[0]
                payload = json.dumps({"event": event, "data": data}, ensure_ascii=False)
                # 安全转义单引号，避免注入
                safe = payload.replace("\\", "\\\\").replace("'", "\\'")
                win.evaluate_js(f"window.__onPyEvent__(JSON.parse('{safe}'))")
        except Exception:
            pass


# ── 启动入口 ────────────────────────────────────────────────────────────

def main():
    try:
        import webview
    except ImportError:
        print("[错误] 缺少 pywebview，请先安装: pip install pywebview")
        sys.exit(1)

    html_path = (WEBVIEW_DIR / "main.html").resolve().as_uri()
    api = Api()

    window = webview.create_window(
        f"Invoice Renamer {APP_VERSION} by {APP_AUTHOR}",
        html_path,
        width=1280,
        height=760,
        min_size=(1040, 660),
        resizable=True,
        js_api=api,
    )
    api.window = window

    def on_loaded():
        try:
            # 直接触发前端初始化（JS 内部有去重保护）
            window.evaluate_js("if (typeof init === 'function') { init(); }")
        except Exception:
            pass

    window.events.loaded += on_loaded
    webview.start(debug=False)


if __name__ == "__main__":
    main()
