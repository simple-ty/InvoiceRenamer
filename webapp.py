"""Invoice Renamer — WebView 版主界面入口（HTTP Server 方案）

彻底放弃 pywebview JS bridge（打包后不稳定），
改用标准库 HTTP server + fetch 通信：
  - 前端 → 后端：fetch('/api/xxx')
  - 后端 → 前端：轮询 /api/poll 或 evaluate_js
  - 文件对话框：pywebview create_file_dialog（Python 端调用）
"""

import json
import os
import sys
import threading
import traceback
from http.server import HTTPServer, BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

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
from excel_exporter import save_invoice_excel, generate_default_filename


def resource_path(relative_path: str) -> str:
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)


# ── 业务 API（与 UI 框架无关）────────────────────────────────────────────

class Api:
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
        self._event_queue: list[dict] = []
        self._event_lock = threading.Lock()
        self._load_cloud_ocr_state()

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

    def _emit(self, event: str, data: dict) -> None:
        """把事件放入队列，前端通过 /api/poll 取走。"""
        self._event_lock.acquire()
        try:
            self._event_queue.append({"event": event, "data": data})
        finally:
            self._event_lock.release()

    def poll_events(self) -> dict:
        self._event_lock.acquire()
        try:
            events = self._event_queue[:]
            self._event_queue.clear()
        finally:
            self._event_lock.release()
        return {"events": events}

    # ── 状态 ──────────────────────────────────────────────────────────

    def get_init_state(self) -> dict:
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
        failed = sum(1 for r in self.records if r.get("status") in ("failed", "cloud_error"))
        not_invoice = sum(1 for r in self.records if r.get("status") == "not_invoice")
        cloud_nc = sum(1 for r in self.records if r.get("status") == "cloud_not_configured")
        return {"total": total, "complete": complete, "partial": partial,
                "failed": failed, "not_invoice": not_invoice,
                "cloud_not_configured": cloud_nc}

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

    # ── 文件选择 ──────────────────────────────────────────────────────

    def choose_folder(self) -> dict:
        try:
            import webview
            win = self.window or (webview.windows[0] if webview.windows else None)
            if not win:
                return {"ok": False, "error": "窗口未就绪"}
            folder_dlg = webview.FileDialog.FOLDER if hasattr(webview, "FileDialog") else webview.FOLDER_DIALOG
            result = win.create_file_dialog(folder_dlg)
            if not result:
                return {"ok": False}
            folder = result[0] if isinstance(result, (list, tuple)) else result
            self.selected_paths = []
            for root_dir, _, files in os.walk(folder):
                for name in files:
                    path = os.path.join(root_dir, name)
                    if path.lower().endswith(ALLOWED_EXTENSIONS):
                        self.selected_paths.append(path)
            return {"ok": True, "path": folder,
                    "summary": f"{len(self.selected_paths)} 个文件"}
        except Exception as e:
            traceback.print_exc()
            return {"ok": False, "error": str(e)}

    def choose_files(self) -> dict:
        try:
            import webview
            win = self.window or (webview.windows[0] if webview.windows else None)
            if not win:
                return {"ok": False, "error": "窗口未就绪"}
            open_dlg = webview.FileDialog.OPEN if hasattr(webview, "FileDialog") else webview.OPEN_DIALOG
            result = win.create_file_dialog(
                open_dlg,
                allow_multiple=True,
                file_types=(
                    "发票文件 (*.pdf;*.jpg;*.jpeg;*.png;*.bmp;*.tiff)",
                    "PDF (*.pdf)",
                    "图片 (*.jpg;*.jpeg;*.png;*.bmp;*.tiff)",
                ),
            )
            if not result:
                return {"ok": False}
            files = list(result) if isinstance(result, (list, tuple)) else [result]
            self.selected_paths = files
            return {"ok": True, "path": f"已选择 {len(files)} 个文件"}
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

    # ── 扫描 ──────────────────────────────────────────────────────────

    def scan_files(self) -> dict:
        if self._scanning or self.processing:
            return {"ok": False, "error": "正在处理中"}
        paths = [p for p in self.selected_paths
                 if os.path.isfile(p) and p.lower().endswith(ALLOWED_EXTENSIONS)]
        if not paths:
            self._emit("status", {"message": "请先选择文件或文件夹"})
            return {"ok": False, "error": "无有效文件"}

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
                                path, self._cloud_secret_id, self._cloud_secret_key)
                        else:
                            parsed = {"fields": {}, "error": "云端未配置", "not_invoice": False,
                                      "_status": "cloud_not_configured"}
                    else:
                        parsed = {"fields": {}, "error": "不支持的类型", "not_invoice": True}

                    fields = parsed.get("fields", {})
                    err = parsed.get("error", "")
                    not_inv = parsed.get("not_invoice", False)
                    force_status = parsed.get("_status", "")

                    current_name = os.path.basename(path)
                    record = {
                        "path": path, "source_name": current_name,
                        "current_name": current_name, "new_name": current_name,
                        "fields": fields, "status": "", "error": err,
                    }

                    # 状态判定（优先级从高到低）
                    if force_status == "cloud_not_configured":
                        record["status"] = "cloud_not_configured"
                    elif not_inv and ("云端" in err or "API" in err or "额度" in err):
                        record["status"] = "cloud_error"
                    elif not_inv:
                        record["status"] = "not_invoice"
                        record["error"] = err or "非发票"
                    elif err:
                        # 有错误但可能是部分识别
                        if self._has_some_fields(fields):
                            record["status"] = "partial"
                        else:
                            record["status"] = "failed"
                    elif self._has_required_fields(fields):
                        record["status"] = "complete"
                    elif self._has_some_fields(fields):
                        record["status"] = "partial"
                    else:
                        record["status"] = "failed"
                        record["error"] = err or "识别失败"
                    record["new_name"] = self._build_target_name(record)
                    records.append(record)
                except Exception as e:
                    records.append({
                        "path": path, "source_name": os.path.basename(path),
                        "current_name": os.path.basename(path),
                        "new_name": os.path.basename(path),
                        "fields": {}, "status": "failed", "error": str(e),
                    })
                self._emit("scan_progress", {"current": i + 1, "total": len(paths)})

            self.records = records
            self._scanning = False
            self._emit("scan_finished", {
                "records": self._serialize_records(),
                "stats": self._calc_stats(),
                "message": "识别完成",
            })

        threading.Thread(target=worker, daemon=True).start()
        return {"ok": True}

    def _has_required_fields(self, fields: dict) -> bool:
        return bool(fields.get("type")) and (bool(fields.get("amount")) or bool(fields.get("date")))

    def _has_some_fields(self, fields: dict) -> bool:
        return any(fields.get(k) for k in ("date", "type", "number", "buyer", "seller", "amount"))

    def _build_target_name(self, record: dict) -> str:
        # 只有 complete / partial 才拼新名字，其余状态保留原名
        if record.get("status") not in ("complete", "partial"):
            return record["current_name"]
        try:
            ext = os.path.splitext(record["path"])[1].lower() or ".pdf"
            return build_name(record["fields"], self.field_enabled,
                              self.field_order, self.custom_value, ext=ext)
        except Exception:
            return record["current_name"]

    # ── 模板 ──────────────────────────────────────────────────────────

    def update_template(self, params: dict) -> dict:
        self.field_order = list(params.get("field_order", self.field_order))
        self.field_enabled = dict(params.get("field_enabled", self.field_enabled))
        self.custom_value = params.get("custom_value", "")
        for r in self.records:
            if r.get("status") in ("complete", "partial"):
                r["new_name"] = self._build_target_name(r)
        return {"ok": True, "records": self._serialize_records()}

    def set_preview_mode(self, params: dict) -> dict:
        self.preview_mode = bool(params.get("preview_mode", True))
        return {"ok": True, "can_rename": not self.preview_mode and bool(self.records)}

    # ── 重命名 ────────────────────────────────────────────────────────

    def on_rename_button_click(self) -> dict:
        if self.rename_history:
            self.undo_rename()
        else:
            self.start_rename()
        return {"ok": True}

    def start_rename(self) -> None:
        if self.processing or self._scanning or not self.records or self.preview_mode:
            return
        self.processing = True
        self._emit("rename_started", {"total": len(self.records)})

        snapshot = list(enumerate(self.records, start=1))
        total = len(snapshot)

        def worker():
            success = failed = skipped = 0
            errors = []
            history = []
            for idx, record in snapshot:
                if record.get("status") != "complete":
                    skipped += 1
                    self._emit("rename_progress", {
                        "current": idx, "total": total,
                        "success": success, "failed": failed, "skipped": skipped})
                    continue
                source_path = record["path"]
                directory = os.path.dirname(source_path)
                ext = os.path.splitext(source_path)[1].lower() or ".pdf"
                target_name = ensure_unique_name(
                    directory,
                    build_name(record["fields"], self.field_enabled,
                               self.field_order, self.custom_value, ext=ext),
                    record["current_name"])
                target_path = os.path.join(directory, target_name)
                try:
                    if record["current_name"] != target_name:
                        os.rename(source_path, target_path)
                        record["path"] = target_path
                        history.append((source_path, target_path, target_name))
                    record["current_name"] = target_name
                    record["new_name"] = target_name
                    success += 1
                    self._emit("rename_item_done", {"idx": idx, "new_name": target_name, "status": "success"})
                except Exception as e:
                    errors.append(f"文件: {record['source_name']}\n错误: {e}\n")
                    failed += 1
                    self._emit("rename_item_done", {"idx": idx, "new_name": record["current_name"], "status": "failed", "error": str(e)[:30]})
                self._emit("rename_progress", {
                    "current": idx, "total": total,
                    "success": success, "failed": failed, "skipped": skipped})

            self.rename_history.extend(history)
            self.processing = False
            msg = "重命名完成"
            err_detail = ""
            if errors:
                msg = f"重命名部分失败，{failed} 个文件出错"
                err_detail = msg
            self._emit("rename_finished", {
                "records": self._serialize_records(),
                "stats": self._calc_stats(),
                "message": msg, "can_undo": bool(self.rename_history),
                "error_detail": err_detail,
            })

        threading.Thread(target=worker, daemon=True).start()

    def undo_rename(self) -> None:
        if not self.rename_history or self.processing or self._scanning:
            return
        self.processing = True
        self._emit("undo_started", {})

        def worker():
            success = failed = 0
            for source_path, target_path, _ in reversed(self.rename_history):
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
                if r.get("status") == "complete":
                    r["new_name"] = self._build_target_name(r)
            self._emit("undo_finished", {
                "records": self._serialize_records(),
                "stats": self._calc_stats(),
                "message": "撤销完成",
                "can_undo": False,
                "error_detail": "撤销部分失败，部分文件已不存在" if failed else "",
            })

        threading.Thread(target=worker, daemon=True).start()

    # ── Excel ────────────────────────────────────────────────────────

    def export_excel(self) -> dict:
        try:
            if not self.records:
                return {"ok": False, "error": "没有可导出的数据"}
            import webview
            win = self.window or (webview.windows[0] if webview.windows else None)
            if not win:
                return {"ok": False, "error": "窗口未就绪"}
            result = win.create_file_dialog(
                webview.FileDialog.SAVE if hasattr(webview, "FileDialog") else webview.SAVE_DIALOG,
                save_filename=generate_default_filename(),
            )
            if not result:
                return {"ok": False, "error": "已取消"}
            path = result if isinstance(result, str) else result[0]
            save_invoice_excel(self.records, path)
            return {"ok": True, "path": path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ── 云端 OCR ──────────────────────────────────────────────────────

    def get_cloud_settings(self) -> dict:
        from cloud_ocr import get_usage_stats
        return {
            "secret_id": self._cloud_secret_id,
            "secret_key": self._cloud_secret_key,
            "enabled": self._cloud_ocr_enabled and self._has_cloud_creds(),
            "usage": get_usage_stats(),
        }

    def save_cloud_settings(self, params: dict) -> dict:
        from cloud_ocr import save_credentials
        sid = params.get("secret_id", "").strip()
        skey = params.get("secret_key", "").strip()
        enabled = bool(params.get("enabled", False))
        save_credentials(sid, skey, enabled=enabled)
        self._cloud_secret_id = sid
        self._cloud_secret_key = skey
        self._cloud_ocr_enabled = enabled
        return {"ok": True, "cloud": {
            "enabled": self._cloud_ocr_enabled and self._has_cloud_creds(),
            "configured": self._has_cloud_creds(), "secret_id": sid}}

    def clear_cloud_settings(self) -> dict:
        from cloud_ocr import clear_credentials
        clear_credentials()
        self._cloud_secret_id = ""
        self._cloud_secret_key = ""
        self._cloud_ocr_enabled = False
        return {"ok": True, "cloud": {"enabled": False, "configured": False, "secret_id": ""}}

    def verify_cloud_credentials(self, params: dict) -> dict:
        from cloud_ocr import validate_credentials
        valid, msg = validate_credentials(
            params.get("secret_id", "").strip(),
            params.get("secret_key", "").strip())
        return {"ok": valid, "message": msg}

    def toggle_cloud_enabled(self) -> dict:
        if not self._has_cloud_creds():
            return {"ok": False, "error": "未配置密钥"}
        self._cloud_ocr_enabled = not self._cloud_ocr_enabled
        from cloud_ocr import save_credentials
        save_credentials(self._cloud_secret_id, self._cloud_secret_key,
                         enabled=self._cloud_ocr_enabled)
        return {"ok": True, "cloud": {
            "enabled": self._cloud_ocr_enabled, "configured": True,
            "secret_id": self._cloud_secret_id}}

    def open_browser(self, params: dict) -> dict:
        try:
            os.startfile(params.get("url", ""))
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}


# ── HTTP 服务器 ──────────────────────────────────────────────────────────

class RequestHandler(BaseHTTPRequestHandler):
    api: Api = None

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/" or path == "/index.html":
            self._serve_static("webview/main.html", "text/html; charset=utf-8")
        elif path == "/css/main.css":
            self._serve_static("webview/css/main.css", "text/css; charset=utf-8")
        elif path == "/js/main.js":
            self._serve_static("webview/js/main.js", "application/javascript; charset=utf-8")
        elif path == "/api/poll":
            self._json(self.api.poll_events())
        elif path == "/api/get_init_state":
            self._json(self.api.get_init_state())
        elif path == "/api/get_cloud_settings":
            self._json(self.api.get_cloud_settings())
        elif path == "/api/clear_cloud_settings":
            self._json(self.api.clear_cloud_settings())
        elif path == "/api/export_excel":
            self._json(self.api.export_excel())
        elif path == "/api/clear_source":
            self._json(self.api.clear_source())
        else:
            self.send_error(404)

    def do_POST(self):
        path = urlparse(self.path).path
        body = b"{}"
        cl = int(self.headers.get("Content-Length", 0))
        if cl > 0:
            body = self.rfile.read(cl)
        try:
            params = json.loads(body) if body else {}
        except Exception:
            params = {}

        if path == "/api/choose_folder":
            self._json(self.api.choose_folder())
        elif path == "/api/choose_files":
            self._json(self.api.choose_files())
        elif path == "/api/scan_files":
            self._json(self.api.scan_files())
        elif path == "/api/update_template":
            self._json(self.api.update_template(params))
        elif path == "/api/toggle_cloud_enabled":
            self._json(self.api.toggle_cloud_enabled())
        elif path == "/api/on_rename_button_click":
            self._json(self.api.on_rename_button_click())
        elif path == "/api/set_preview_mode":
            self._json(self.api.set_preview_mode(params))
        elif path == "/api/save_cloud_settings":
            self._json(self.api.save_cloud_settings(params))
        elif path == "/api/verify_cloud_credentials":
            self._json(self.api.verify_cloud_credentials(params))
        elif path == "/api/open_browser":
            self._json(self.api.open_browser(params))
        else:
            self.send_error(404)

    def _json(self, data):
        b = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

    def _serve_static(self, relative_path, content_type):
        filepath = resource_path(relative_path)
        try:
            with open(filepath, "rb") as f:
                content = f.read()
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
        except Exception as e:
            self.send_error(404, str(e))

    def log_message(self, *args):
        pass


def find_free_port() -> int:
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


# ── 启动入口 ────────────────────────────────────────────────────────────

def main():
    try:
        import webview
    except ImportError:
        print("[错误] 缺少 pywebview，请先安装: pip install pywebview")
        sys.exit(1)

    api = Api()
    RequestHandler.api = api

    port = find_free_port()
    server = ThreadingHTTPServer(("127.0.0.1", port), RequestHandler)
    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()

    url = f"http://127.0.0.1:{port}/"
    window = webview.create_window(
        f"Invoice Renamer {APP_VERSION} by {APP_AUTHOR}",
        url,
        width=1280,
        height=760,
        min_size=(1040, 660),
        resizable=True,
    )
    api.window = window
    webview.start(debug=False)
    server.shutdown()


if __name__ == "__main__":
    main()
