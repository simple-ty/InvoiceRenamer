"""Invoice Renamer — 主界面模块

依赖自动安装逻辑保留在此文件顶部，确保打包后仍能自检依赖。
"""

import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

# ── 依赖自检与自动安装 ──────────────────────────────────────────────────
_REQUIRED_PACKAGES = [
    ("pdfplumber", "pdfplumber"),
    ("customtkinter", "customtkinter"),
    ("openpyxl", "openpyxl"),
]

for _module_name, _package_name in _REQUIRED_PACKAGES:
    try:
        __import__(_module_name)
    except ImportError:
        if getattr(sys, "frozen", False):
            print(f"[依赖缺失] 打包程序缺少依赖 {_module_name}，请重新打包。")
            sys.exit(1)
        print("-" * 50)
        print(f"[安装] 缺少依赖 {_module_name}，正在自动安装...")
        print(f"[安装] 执行: pip install {_package_name}")
        print("[安装] 首次安装需要联网下载，请稍等...\n")
        sys.stdout.flush()
        try:
            _result = subprocess.run(
                [sys.executable, "-m", "pip", "install", _package_name, "--user"],
                capture_output=False,
                timeout=180,
                text=True,
            )
            if _result.returncode == 0:
                print(f"\n[安装] {_module_name} 安装成功!\n")
            else:
                print(f"\n[安装失败] 退出码: {_result.returncode}")
                print(f"[安装失败] 请手动执行: pip install {_package_name}")
                sys.stdout.flush()
                sys.exit(1)
        except subprocess.TimeoutExpired:
            print(f"\n[安装超时] pip install {_package_name} 超过 180 秒")
            print(f"[安装超时] 请手动执行: pip install {_package_name}")
            sys.stdout.flush()
            sys.exit(1)
        except Exception as exc:
            print(f"\n[安装出错] {exc}")
            print(f"[安装出错] 请手动执行: pip install {_package_name}")
            sys.stdout.flush()
            sys.exit(1)
    except Exception as exc:
        print(f"\n[检查依赖出错] {exc}")
        sys.exit(1)

# ── 第三方库 ─────────────────────────────────────────────────────────────
import customtkinter as ctk

# ── 项目模块 ─────────────────────────────────────────────────────────────
from config import (  # noqa: E402
    APP_AUTHOR, APP_ID, APP_VERSION,
    DEFAULT_FIELD_ENABLED, DEFAULT_FIELD_ORDER,
    FIELD_LABELS, RESULT_HINT,
    IMAGE_EXTENSIONS, ALLOWED_EXTENSIONS,
    STAT_CARD_STYLE, TREE_TAG_COLORS,
)
from invoice_parser import parse_invoice, parse_image_cloud  # noqa: E402
from name_builder import build_name, ensure_unique_name, sanitize_part  # noqa: E402
from excel_exporter import export_invoice_excel  # noqa: E402


# ── 工具函数 ─────────────────────────────────────────────────────────────

def resource_path(relative_path: str) -> str:
    """获取打包后的资源文件路径（PyInstaller 兼容）。"""
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)


def setup_windows_app_id() -> None:
    """设置 Windows 任务栏图标应用 ID（仅 Windows）。"""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception:
        pass


# ── 主应用类 ────────────────────────────────────────────────────────────

class App:
    def __init__(self, root: ctk.CTk):
        self.root = root
        self.root.title(f"Invoice Renamer {APP_VERSION} by {APP_AUTHOR}")
        self.root.geometry("1280x760")
        self.root.minsize(1040, 660)
        self.root.configure(fg_color="#F5F5F5")

        self.selected_paths: list[str] = []
        self.records: list[dict] = []
        self.processing: bool = False
        self._scanning: bool = False
        self.rename_history: list[tuple[str, str, str]] = []  # (source, target, new_name)
        self.field_order: list[str] = list(DEFAULT_FIELD_ORDER)

        # StringVar / BooleanVar
        self.path_var = tk.StringVar()
        self.summary_var = tk.StringVar(value=RESULT_HINT)
        self.status_var = tk.StringVar(value="就绪")
        self.custom_value_var = tk.StringVar()
        self.preview_mode_var = tk.BooleanVar(value=True)
        self.template_vars = {
            k: tk.BooleanVar(value=DEFAULT_FIELD_ENABLED[k])
            for k in DEFAULT_FIELD_ORDER
        }
        self.stats_vars = {
            k: tk.StringVar(value="0") for k in STAT_CARD_STYLE
        }

        self.custom_value_var.trace_add("write", lambda *_: self.on_template_change())

        # 云端 OCR 状态
        self._cloud_ocr_enabled = False
        self._cloud_secret_id = ""
        self._cloud_secret_key = ""
        self._load_cloud_ocr_state()

        self._prepare_appearance()
        self._set_icon()
        self._configure_treeview_style()
        self._build_ui()
        self.refresh_template_panel()
        self.load_from_argv()

    # ── 初始化辅助 ──────────────────────────────────────────────────────

    def _prepare_appearance(self) -> None:
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("green")
        yahei = "Microsoft YaHei UI"
        self.font_title = ctk.CTkFont(family=yahei, size=20, weight="bold")
        self.font_subtitle = ctk.CTkFont(family=yahei, size=13)
        self.font_section = ctk.CTkFont(family=yahei, size=15, weight="bold")
        self.font_card_value = ctk.CTkFont(family=yahei, size=18, weight="bold")
        self.font_container = ctk.CTkFont(family=yahei, size=12)
        self.font_button = ctk.CTkFont(family=yahei, size=12, weight="bold")
        self.font_small = ctk.CTkFont(family=yahei, size=10)
        # 弹窗/卡片专用字体（字号稍大，阅读更舒适）
        self.font_body = ctk.CTkFont(family=yahei, size=14)
        self.font_hint = ctk.CTkFont(family=yahei, size=12)

        # 云端 OCR 图标：固定天蓝色，不随状态变化
        self._cloud_icon = self._create_cloud_icon("#3FA9F5", 30)

    def _set_icon(self) -> None:
        try:
            self.root.iconbitmap(resource_path("icon.ico"))
        except Exception:
            pass

    def _configure_treeview_style(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "Invoice.Treeview",
            rowheight=28,
            font=("Microsoft YaHei UI", 9),
            background="#FFFFFF",
            foreground="#191919",
            fieldbackground="#FFFFFF",
            borderwidth=0,
        )
        style.configure(
            "Invoice.Treeview.Heading",
            font=("Microsoft YaHei UI", 9, "bold"),
            background="#F2F3F5",
            foreground="#191919",
            borderwidth=0,
            relief="flat",
        )
        style.map(
            "Invoice.Treeview",
            background=[("selected", "#D8F5E5")],
            foreground=[("selected", "#191919")],
        )
        style.layout("Invoice.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])

    def _build_ui(self) -> None:
        outer = ctk.CTkFrame(self.root, fg_color="transparent")
        outer.pack(fill="both", expand=True, padx=12, pady=10)

        self.build_header(outer)

        content = ctk.CTkFrame(outer, fg_color="transparent")
        content.pack(fill="both", expand=True, pady=(8, 0))
        content.grid_columnconfigure(1, weight=1)
        content.grid_rowconfigure(0, weight=1)

        self.template_panel = ctk.CTkFrame(content, fg_color="#FFFFFF", corner_radius=8, width=250)
        self.template_panel.grid(row=0, column=0, sticky="nsw", padx=(0, 10))
        self.template_panel.grid_propagate(False)
        self.template_panel.pack_propagate(False)

        self.results_panel = ctk.CTkFrame(content, fg_color="#FFFFFF", corner_radius=8)
        self.results_panel.grid(row=0, column=1, sticky="nsew")
        self.results_panel.grid_columnconfigure(0, weight=1)
        self.results_panel.grid_rowconfigure(1, weight=1)

        self.build_template_panel()
        self.build_results_panel()
        self.build_action_bar(outer)

    # ── Header ─────────────────────────────────────────────────────────

    def build_header(self, parent: ctk.CTkFrame) -> None:
        frame = ctk.CTkFrame(parent, fg_color="#FFFFFF", corner_radius=8)
        frame.pack(fill="x")
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=0)

        left = ctk.CTkFrame(frame, fg_color="transparent")
        left.grid(row=0, column=0, sticky="w", padx=16, pady=10)
        title_line = ctk.CTkFrame(left, fg_color="transparent")
        title_line.pack(anchor="w")
        ctk.CTkLabel(
            title_line, text="Invoice Renamer", font=self.font_title, text_color="#191919"
        ).pack(side="left")
        ctk.CTkLabel(
            title_line,
            text="批量识别、预览、导出、重命名。现在支持可排序模板字段和自定义字段。",
            font=self.font_subtitle,
            text_color="#4C4C4C",
        ).pack(side="left", padx=(14, 0))

        source_area = ctk.CTkFrame(frame, fg_color="#F7F7F7", corner_radius=8)
        source_area.grid(row=0, column=1, sticky="e", padx=16, pady=10)
        source_area.grid_columnconfigure(0, weight=1)

        self.path_entry = ctk.CTkEntry(
            source_area,
            textvariable=self.path_var,
            placeholder_text="请选择文件夹，或选择多个 PDF 文件",
            width=250,
            height=32,
            corner_radius=6,
            font=self.font_container,
            fg_color="#FFFFFF",
            border_color="#E5E5E5",
        )
        self.path_entry.grid(row=0, column=0, sticky="ew", padx=(8, 8), pady=8)

        self._source_btn = ctk.CTkButton(
            source_area, text="选择来源 ▾", command=self._popup_source_menu,
            width=100, height=32, corner_radius=6,
            fg_color="#07C160", hover_color="#06AD56", text_color="#191919",
            font=self.font_button,
        )
        self._source_btn.grid(row=0, column=1, padx=(0, 8), pady=8)

        frame.bind("<Configure>", self.update_path_entry_width, add="+")
        self.root.after_idle(self.update_path_entry_width)

    def update_path_entry_width(self, *_):
        if not hasattr(self, "path_entry"):
            return
        width = max(280, min(460, int(self.root.winfo_width() * 0.28)))
        self.path_entry.configure(width=width)

    # ── 统计卡片 ────────────────────────────────────────────────────────
    def build_result_stats(self, parent: ctk.CTkFrame) -> None:
        stats_area = ctk.CTkFrame(parent, fg_color="transparent")
        stats_area.grid(row=0, column=2, sticky="e")
        for idx, key in enumerate(STAT_CARD_STYLE):
            w = 112 if key == "failed" else 94
            chip = ctk.CTkFrame(
                stats_area, fg_color="#F7F7F7", corner_radius=6, width=w, height=30
            )
            chip.grid(row=0, column=idx, padx=(0 if idx == 0 else 6, 0))
            chip.grid_propagate(False)
            ctk.CTkLabel(
                chip, text=STAT_CARD_STYLE[key]["title"],
                font=self.font_container, text_color="#4C4C4C",
            ).pack(side="left", padx=(8, 4), pady=5)
            ctk.CTkLabel(
                chip, textvariable=self.stats_vars[key],
                font=self.font_button,
                text_color=STAT_CARD_STYLE[key]["accent"],
            ).pack(side="right", padx=(4, 8), pady=5)

    # ── 模板面板 ────────────────────────────────────────────────────────

    def _create_cloud_icon(self, color: str = "#07C160", size: int = 20) -> ctk.CTkImage:
        """生成一个实心 iCloud 风格云朵图标。"""
        try:
            from PIL import Image, ImageDraw

            def _bezier3(p0, p1, p2, p3, t):
                u = 1 - t
                return (
                    u**3 * p0[0] + 3 * u**2 * t * p1[0] + 3 * u * t**2 * p2[0] + t**3 * p3[0],
                    u**3 * p0[1] + 3 * u**2 * t * p1[1] + 3 * u * t**2 * p2[1] + t**3 * p3[1],
                )

            # iCloud 风格云朵轮廓（viewBox 80x54）
            segments = [
                ("M", (55.0, 46.0)),
                ("H", (20.0, 46.0)),
                ("C", (9.5, 46.0), (2.0, 38.5), (2.0, 30.0)),
                ("C", (2.0, 22.5), (7.5, 16.5), (15.0, 16.0)),
                ("C", (16.5, 8.5), (23.5, 3.0), (32.0, 3.0)),
                ("C", (39.0, 3.0), (45.0, 7.0), (47.5, 13.0)),
                ("C", (50.0, 12.0), (53.0, 11.0), (56.0, 11.0)),
                ("C", (66.5, 11.0), (74.0, 19.0), (74.0, 29.0)),
                ("C", (74.0, 38.5), (66.5, 46.0), (55.0, 46.0)),
            ]

            points = []
            current = None
            for seg in segments:
                cmd = seg[0]
                if cmd == "M":
                    current = seg[1]
                    points.append(current)
                elif cmd == "H":
                    # 水平线也按曲线处理以保持圆滑：用当前点和终点构造一个退化三次贝塞尔
                    p0, p3 = current, (seg[1][0], current[1])
                    p1 = (p0[0] * 0.67 + p3[0] * 0.33, p0[1])
                    p2 = (p0[0] * 0.33 + p3[0] * 0.67, p3[1])
                    for i in range(1, 21):
                        pt = _bezier3(p0, p1, p2, p3, i / 20.0)
                        points.append(pt)
                    current = p3
                elif cmd == "C":
                    p0 = current
                    p1, p2, p3 = seg[1], seg[2], seg[3]
                    for i in range(1, 21):
                        pt = _bezier3(p0, p1, p2, p3, i / 20.0)
                        points.append(pt)
                    current = p3

            # 缩放并居中到 size x size 画布
            xs = [p[0] for p in points]
            ys = [p[1] for p in points]
            min_x, max_x = min(xs), max(xs)
            min_y, max_y = min(ys), max(ys)
            scale = min((size - 2) / (max_x - min_x), (size - 2) / (max_y - min_y))
            off_x = (size - (max_x - min_x) * scale) / 2
            off_y = (size - (max_y - min_y) * scale) / 2
            scaled = [
                (off_x + (p[0] - min_x) * scale, off_y + (p[1] - min_y) * scale)
                for p in points
            ]

            img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            c = tuple(int(color[i:i + 2], 16) for i in (1, 3, 5)) + (255,)
            draw.polygon(scaled, fill=c)
            return ctk.CTkImage(light_image=img, dark_image=img, size=(size, size))
        except Exception:
            return None

    def build_template_panel(self) -> None:
        pnl = self.template_panel

        # 标题
        ctk.CTkLabel(
            pnl, text="命名模板", font=self.font_section, text_color="#191919"
        ).pack(anchor="w", padx=12, pady=(12, 2))
        ctk.CTkLabel(
            pnl,
            text="勾选需要的字段，拖拽手柄（≡）调整顺序。自定义字段可用于项目名、部门名、报销批次等固定内容。",
            font=self.font_subtitle, text_color="#4C4C4C",
            justify="left", wraplength=214,
        ).pack(anchor="w", padx=12)

        # 中间内容区（撑开剩余空间）
        middle_frame = ctk.CTkFrame(pnl, fg_color="transparent")
        middle_frame.pack(fill="both", expand=True, padx=0, pady=0)

        self.template_rows_frame = ctk.CTkScrollableFrame(
            middle_frame, fg_color="#F7F7F7", corner_radius=6, border_width=0, height=266,
        )
        self.template_rows_frame.pack(fill="x", expand=False, padx=10, pady=(8, 4))
        self.template_rows_frame.bind("<Configure>", self.update_template_scrollbar_visibility, add="+")
        self.template_rows_frame._parent_canvas.bind("<Configure>", self.update_template_scrollbar_visibility, add="+")

        # 自定义字段输入
        custom_box = ctk.CTkFrame(middle_frame, fg_color="#F7F7F7", corner_radius=6)
        custom_box.pack(fill="x", padx=10, pady=(0, 0))
        ctk.CTkLabel(
            custom_box, text="自定义字段内容", font=self.font_container, text_color="#191919"
        ).pack(anchor="w", padx=10, pady=(8, 4))
        self.custom_entry = ctk.CTkEntry(
            custom_box, textvariable=self.custom_value_var,
            placeholder_text="例如：项目A / 2026Q2 报销",
            height=32, corner_radius=6,
            font=self.font_container, fg_color="#FFFFFF", border_color="#E5E5E5",
        )
        self.custom_entry.pack(fill="x", padx=10, pady=(0, 5))
        ctk.CTkLabel(
            custom_box,
            text='只有在勾选"自定义字段"后，输入的内容才会参与命名。',
            font=self.font_subtitle, text_color="#4C4C4C",
            justify="left", wraplength=210,
        ).pack(anchor="w", padx=10, pady=(0, 8))

        # 底部：云端 OCR 设置入口（整行可点击、整行悬浮）
        cloud_row = ctk.CTkFrame(
            pnl, fg_color="#FFFFFF", corner_radius=8,
            border_width=1, border_color="#E5E5E5", height=54,
        )
        cloud_row.pack(fill="x", side="bottom", padx=10, pady=(0, 8))
        cloud_row.pack_propagate(False)
        cloud_row.configure(cursor="hand2")

        self._cloud_hover = False

        def _on_cloud_enter(_):
            self._cloud_hover = True
            cloud_row.configure(fg_color="#F2F7F2")
        def _on_cloud_leave(_):
            self._cloud_hover = False
            cloud_row.after(10, lambda: _check_cloud_leave())
        def _check_cloud_leave():
            if not self._cloud_hover:
                cloud_row.configure(fg_color="#FFFFFF")

        cloud_row.bind("<Enter>", _on_cloud_enter)
        cloud_row.bind("<Leave>", _on_cloud_leave)

        def _open_cloud_settings(_=None):
            self._open_cloud_ocr_settings()
        cloud_row.bind("<Button-1>", _open_cloud_settings)

        # 左侧图标 + 文字
        self._cloud_ocr_btn = ctk.CTkLabel(
            cloud_row,
            text="  云端识别设置",
            image=self._cloud_icon,
            compound="left",
            font=self.font_body,
            text_color="#191919",
        )
        self._cloud_ocr_btn.pack(side="left", padx=(10, 0))
        self._cloud_ocr_btn.bind("<Button-1>", _open_cloud_settings)
        self._cloud_ocr_btn.bind("<Enter>", _on_cloud_enter)
        self._cloud_ocr_btn.bind("<Leave>", _on_cloud_leave)

        # 右侧状态按钮：已配置时可点击切换开关，未配置时显示红色● 未配置
        def _toggle_cloud_state():
            if not self._has_cloud_creds():
                return
            self._cloud_ocr_enabled = not self._cloud_ocr_enabled
            from cloud_ocr import save_credentials
            save_credentials(self._cloud_secret_id, self._cloud_secret_key, enabled=self._cloud_ocr_enabled)
            self._update_cloud_status_dot()
            self.update_rename_button_state()

        self._cloud_status_btn = ctk.CTkButton(
            cloud_row,
            text="● 未配置",
            font=self.font_hint,
            text_color="#FA5151",
            fg_color="#FCEBEB",
            hover_color="#FCEBEB",
            corner_radius=10,
            width=76,
            height=24,
            cursor="arrow",
            command=_toggle_cloud_state,
        )
        self._cloud_status_btn.pack(side="right", padx=(0, 12))
        self._cloud_status_btn.bind("<Enter>", _on_cloud_enter)
        self._cloud_status_btn.bind("<Leave>", _on_cloud_leave)

        self._update_cloud_status_dot()

    def refresh_template_panel(self) -> None:
        """刷新命名模板面板，支持拖拽排序。"""
        for child in self.template_rows_frame.winfo_children():
            child.destroy()
        self._row_frames: list[ctk.CTkFrame] = []
        self._row_badges: list[ctk.CTkLabel] = []

        for idx, key in enumerate(self.field_order):
            card = ctk.CTkFrame(self.template_rows_frame, fg_color="#FFFFFF", corner_radius=6)
            card.pack(fill="x", pady=(0, 3), padx=2)
            self._row_frames.append(card)

            badge = ctk.CTkLabel(
                card, text=f"{idx + 1:02d}", width=28, height=24,
                corner_radius=6, fg_color="#F2F3F5",
                text_color="#576B95", font=self.font_small,
            )
            badge.pack(side="left", padx=(6, 6), pady=3)
            self._row_badges.append(badge)

            checkbox = ctk.CTkCheckBox(
                card, text="", width=20,
                variable=self.template_vars[key],
                command=self.on_template_change,
                fg_color="#07C160", hover_color="#06AD56",
                checkmark_color="#F9F8F4",
            )
            checkbox.pack(side="left")

            text_frame = ctk.CTkFrame(card, fg_color="transparent")
            text_frame.pack(side="left", fill="x", expand=True, padx=(6, 4), pady=3)
            ctk.CTkLabel(
                text_frame, text=FIELD_LABELS[key],
                font=self.font_container, text_color="#191919",
            ).pack(anchor="w")

            # 拖拽手柄放在右侧，用三道横线图标
            handle = ctk.CTkLabel(
                card, text="\u2261", width=32, height=28,
                corner_radius=6, fg_color="#F2F3F5",
                text_color="#AAAAAA", font=("Consolas", 20),
                cursor="hand2",
            )
            handle.pack(side="right", padx=(4, 8), pady=3)
            # 绑定 frame 引用而非固定索引，避免重排后索引错位
            handle.bind("<ButtonPress-1>", lambda e, f=card: self._on_drag_start(e, f))

        self.root.after_idle(self.update_template_scrollbar_visibility)

    def _on_drag_start(self, event, frame) -> None:
        """开始拖拽：通过 frame 引用查找当前索引，绑定全局鼠标事件。"""
        if frame not in self._row_frames:
            return
        self._drag_index = self._row_frames.index(frame)
        if 0 <= self._drag_index < len(self._row_frames):
            self._row_frames[self._drag_index].configure(fg_color="#E8F5E9")
        # 绑定到 root，确保鼠标移出手柄区域仍能跟踪
        self.root.bind("<B1-Motion>", self._on_drag_motion)
        self.root.bind("<ButtonRelease-1>", self._on_drag_release)

    def _get_row_at_y(self, y_root: int) -> "int | None":
        """根据鼠标全局 Y 坐标判断指向哪一行"""
        for i, frame in enumerate(self._row_frames):
            fy = frame.winfo_rooty()
            fh = frame.winfo_height()
            if fy <= y_root <= fy + fh:
                return i
        # 鼠标在所有行上方 → 第一行
        if self._row_frames:
            top = self._row_frames[0].winfo_rooty()
            if y_root < top:
                return 0
        return None

    def _on_drag_motion(self, event) -> None:
        """拖拽过程中：实时调整排序，鼠标到哪里行就移到哪里。
        不销毁/重建控件，只重排现有帧，避免丢失事件绑定。"""
        if not hasattr(self, "_drag_index"):
            return
        target_idx = self._get_row_at_y(event.y_root)
        if target_idx is None or target_idx == self._drag_index:
            return

        # 更新 field_order
        key = self.field_order.pop(self._drag_index)
        self.field_order.insert(target_idx, key)

        # 重排 _row_frames 和 _row_badges 列表（不销毁控件）
        frame = self._row_frames.pop(self._drag_index)
        self._row_frames.insert(target_idx, frame)
        badge = self._row_badges.pop(self._drag_index)
        self._row_badges.insert(target_idx, badge)

        # 重新 pack 所有帧到新顺序
        for f in self._row_frames:
            f.pack_forget()
        for f in self._row_frames:
            f.pack(fill="x", pady=(0, 3), padx=2)

        # 同步更新所有徽标序号
        for i, b in enumerate(self._row_badges):
            b.configure(text=f"{i + 1:02d}")

        self._drag_index = target_idx

        # 高亮当前拖拽行
        for f in self._row_frames:
            f.configure(fg_color="#FFFFFF")
        if 0 <= self._drag_index < len(self._row_frames):
            self._row_frames[self._drag_index].configure(fg_color="#C8E6C9")

    def _on_drag_release(self, event) -> None:
        """释放鼠标：清理拖拽状态。"""
        if not hasattr(self, "_drag_index"):
            return
        # 清除高亮
        for frame in self._row_frames:
            try:
                frame.configure(fg_color="#FFFFFF")
            except Exception:
                pass
        del self._drag_index
        # 解绑 root 上的全局拖拽事件
        self.root.unbind("<B1-Motion>")
        self.root.unbind("<ButtonRelease-1>")
        # 排序已在拖拽过程中实时完成，这里只需触发命名刷新
        self.on_template_change()

    def update_template_scrollbar_visibility(self, *_):
        frame = getattr(self, "template_rows_frame", None)
        if frame is None:
            return
        canvas = getattr(frame, "_parent_canvas", None)
        scrollbar = getattr(frame, "_scrollbar", None)
        if canvas is None or scrollbar is None:
            return
        bbox = canvas.bbox("all")
        if not bbox:
            return
        content_height = bbox[3] - bbox[1]
        canvas_height = canvas.winfo_height()
        if content_height <= canvas_height + 2:
            scrollbar.grid_remove()
        else:
            scrollbar.grid()

    def move_field(self, key: str, delta: int) -> None:
        current = self.field_order.index(key)
        target = current + delta
        if target < 0 or target >= len(self.field_order):
            return
        self.field_order[current], self.field_order[target] = (
            self.field_order[target], self.field_order[current],
        )
        self.refresh_template_panel()
        self.on_template_change()

    # ── 结果表格面板 ────────────────────────────────────────────────────

    def build_results_panel(self) -> None:
        top = ctk.CTkFrame(self.results_panel, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
        top.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            top, text="识别结果", font=self.font_section, text_color="#191919"
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            top, textvariable=self.summary_var,
            font=self.font_subtitle, text_color="#4C4C4C",
            justify="left", wraplength=480,
        ).grid(row=0, column=1, sticky="w", padx=(12, 12))

        self.build_result_stats(top)

        # Treeview
        table_shell = ctk.CTkFrame(
            self.results_panel, fg_color="#FFFFFF",
            corner_radius=6, border_width=1, border_color="#E5E5E5",
        )
        table_shell.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 10))
        table_shell.grid_columnconfigure(0, weight=1)
        table_shell.grid_rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            table_shell,
            columns=("idx", "org", "new", "type", "seller", "amount", "status"),
            show="headings", style="Invoice.Treeview", selectmode="browse", height=12,
        )
        columns_cfg = [
            ("idx", "#", 44, tk.CENTER),
            ("org", "原文件名", 290, tk.W),
            ("new", "新文件名", 380, tk.W),
            ("type", "发票类型", 120, tk.W),
            ("seller", "销售方/行程信息", 210, tk.W),
            ("amount", "金额", 100, tk.CENTER),
            ("status", "状态", 120, tk.CENTER),
        ]
        # 排序状态
        self._sort_col = None
        self._sort_reverse = False
        for key, heading, width, anchor in columns_cfg:
            self.tree.heading(
                key, text=heading, anchor=anchor,
                command=lambda c=key: self._sort_by_column(c),
            )
            self.tree.column(key, width=width, anchor=anchor, minwidth=40,
                             stretch=False)
        for tag, color in TREE_TAG_COLORS.items():
            self.tree.tag_configure(tag, foreground=color)

        scrollbar = ctk.CTkScrollbar(
            table_shell, orientation="vertical", command=self.tree.yview,
            width=12, corner_radius=6,
            fg_color="#F2F3F5", button_color="#C9CDD4", button_hover_color="#AEB4BE",
        )
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=(8, 0))
        scrollbar.grid(row=0, column=1, sticky="ns", padx=(0, 8), pady=(8, 0))

        # 水平滚动条
        h_scrollbar = ctk.CTkScrollbar(
            table_shell, orientation="horizontal", command=self.tree.xview,
            height=12, corner_radius=6,
            fg_color="#F2F3F5", button_color="#C9CDD4", button_hover_color="#AEB4BE",
        )
        self.tree.configure(xscrollcommand=h_scrollbar.set)
        h_scrollbar.grid(row=1, column=0, sticky="ew", padx=(8, 0), pady=(0, 8))

    # ── 底部操作栏 ──────────────────────────────────────────────────────

    def build_action_bar(self, parent: ctk.CTkFrame) -> None:
        bar = ctk.CTkFrame(parent, fg_color="#FFFFFF", corner_radius=8)
        bar.pack(fill="x")
        bar.grid_columnconfigure(1, weight=1)

        self.status_label = ctk.CTkLabel(
            bar, textvariable=self.status_var,
            font=self.font_subtitle, text_color="#4C4C4C",
        )
        self.status_label.grid(row=0, column=0, sticky="w", padx=(12, 10), pady=8)

        self.progress_bar = ctk.CTkProgressBar(
            bar, height=8, corner_radius=4, progress_color="#07C160",
        )
        self.progress_bar.grid(row=0, column=1, sticky="ew", padx=(0, 10), pady=8)
        self.progress_bar.set(0)

        self.preview_switch = ctk.CTkSwitch(
            bar, text="预览模式（改名请关闭）",
            variable=self.preview_mode_var,
            onvalue=True, offvalue=False,
            command=self.update_rename_button_state,
            font=self.font_container, text_color="#191919",
            progress_color="#07C160",
            button_color="#FFFFFF", button_hover_color="#F7F7F7",
        )
        self.preview_switch.grid(row=0, column=2, sticky="e", padx=(0, 12), pady=8)

        action_group = ctk.CTkFrame(bar, fg_color="transparent")
        action_group.grid(row=0, column=3, sticky="e", padx=(0, 12), pady=8)

        ctk.CTkButton(
            action_group, text="重新识别", command=self.scan_files,
            width=92, height=32, corner_radius=6,
            fg_color="#F2F3F5", hover_color="#EDEDED", text_color="#191919",
            font=self.font_container,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            action_group, text="导出 Excel", command=self.export_excel,
            width=92, height=32, corner_radius=6,
            fg_color="#F2F3F5", hover_color="#EDEDED", text_color="#191919",
            font=self.font_container,
        ).pack(side="left", padx=(0, 8))

        self.rename_button = ctk.CTkButton(
            action_group, text="开始重命名", command=self.on_rename_button_click,
            width=120, height=32, corner_radius=6,
            fg_color="#07C160", hover_color="#06AD56", text_color="#191919",
            font=self.font_button,
        )
        self.rename_button.pack(side="left")
        self.update_rename_button_state()
        self._sync_rename_button()

        # 窗口 resize 时主动触发重绘，避免 customtkinter 控件出现黑块
        self._resize_job = None
        self.root.bind("<Configure>", self._on_window_resize)

    def _on_window_resize(self, event) -> None:
        """窗口大小改变时，延迟 150ms 后强制整个窗口重绘，消除 customtkinter 黑块。"""
        if event.widget != self.root:
            return
        if self._resize_job is not None:
            self.root.after_cancel(self._resize_job)
        self._resize_job = self.root.after(150, self._force_redraw)

    def _force_redraw(self) -> None:
        """用 alpha 技巧强制 Windows 重绘整个窗口，彻底消除 customtkinter 黑块。"""
        self._resize_job = None
        # 临时将透明度设为 0.99 再恢复为 1.0，触发 Windows 完整重绘
        # 用户完全察觉不到变化，但能强制所有 customtkinter 控件重绘
        self.root.attributes("-alpha", 0.99)
        self.root.after(10, lambda: self.root.attributes("-alpha", 1.0))

    # ── 合并按钮回调 ──────────────────────────────────────────────────

    def on_rename_button_click(self) -> None:
        """智能按钮点击：有撤销记录时执行撤销，否则执行重命名。"""
        if self.rename_history:
            self.undo_rename()
        else:
            self.start_rename()

    def _sync_rename_button(self) -> None:
        """根据 rename_history 同步按钮文字与颜色。"""
        if not hasattr(self, "rename_button"):
            return
        if self.rename_history:
            self.rename_button.configure(
                text="撤销重命名",
                fg_color="#F9D65C", hover_color="#E8C65C",
                text_color="#191919",
            )
        else:
            self.rename_button.configure(
                text="开始重命名",
                fg_color="#07C160", hover_color="#06AD56",
                text_color="#191919",
            )

    # ── 状态与按钮 ──────────────────────────────────────────────────────

    def set_status(self, message: str) -> None:
        self.status_var.set(message)

    def update_rename_button_state(self) -> None:
        if not hasattr(self, "rename_button"):
            return
        scanning = getattr(self, "_scanning", False)
        state = "disabled" if self.processing or scanning or self.preview_mode_var.get() else "normal"
        self.rename_button.configure(state=state)
        # 同步文字（处理中时不覆盖）
        if not self.processing:
            self._sync_rename_button()

    # ── 重命名执行 ──────────────────────────────────────────────────────

    def start_rename(self) -> None:
        if self.processing or getattr(self, "_scanning", False) or not self.records:
            return
        if self.preview_mode_var.get():
            self.set_status("请关闭预览模式后再执行重命名。")
            self.update_rename_button_state()
            return

        self.processing = True
        self.rename_button.configure(text="处理中...")
        self.update_rename_button_state()

        # 在主线程捕获模板配置，避免后台线程访问 Tkinter 变量
        enabled_map = self.current_enabled_map()
        field_order = list(self.field_order)
        custom_value = self.custom_value_var.get()
        records_snapshot = list(enumerate(self.records, start=1))
        items = self.tree.get_children()
        total = len(records_snapshot)

        def worker() -> None:
            success = failed = skipped = 0
            errors: list[str] = []
            items_snapshot = tuple(items)  # 锁定快照

            for idx, record in records_snapshot:
                if idx > len(items_snapshot):
                    break  # 防御性保护，防止索引越界
                item = items_snapshot[idx - 1]
                if self.is_unrecognized_record(record):
                    self.root.after(0, self._ui_rename_skipped, item, record)
                    skipped += 1
                else:
                    source_path = record["path"]
                    directory = os.path.dirname(source_path)
                    ext = os.path.splitext(source_path)[1].lower() or ".pdf"
                    target_name = ensure_unique_name(
                        directory,
                        build_name(record["fields"], enabled_map, field_order, custom_value, ext=ext),
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
                        self.root.after(0, self._ui_rename_success, item, target_name, source_path, target_path, rename_happened)
                        success += 1
                    except Exception as e:
                        import traceback
                        tb = traceback.format_exc()
                        errors.append(f"文件: {record['source_name']}\n错误: {e}\n{tb}\n")
                        self.root.after(0, self._ui_rename_failed, item, str(e)[:30])
                        failed += 1

                self.root.after(0, self._ui_rename_progress, idx, total, success, failed, skipped)

            self.root.after(0, self._rename_finished, success, failed, skipped, errors)

        threading.Thread(target=worker, daemon=True).start()

    # ── 重命名 UI 回调（均在主线程执行）────────────────────────────

    def _ui_rename_skipped(self, item, record) -> None:
        self.tree.set(item, "new", record["current_name"])
        self.tree.set(item, "status", record["error"] or "未识别")
        self.tree.item(item, tags=("error" if record["error"] else "idle",))

    def _ui_rename_success(self, item, target_name, source_path, target_path, rename_happened) -> None:
        self.tree.set(item, "new", target_name)
        self.tree.set(item, "status", "完成")
        self.tree.item(item, tags=("success",))
        if rename_happened:
            self.rename_history.append((source_path, target_path, target_name))

    def _ui_rename_failed(self, item, error_msg) -> None:
        self.tree.set(item, "status", f"失败:{error_msg}")
        self.tree.item(item, tags=("error",))

    def _ui_rename_progress(self, idx: int, total: int, success: int, failed: int, skipped: int) -> None:
        self.progress_bar.set(idx / total)
        self.set_status(f"重命名中... ({idx}/{total})")

    def _rename_finished(self, success: int, failed: int, skipped: int, errors: list[str]) -> None:
        self.progress_bar.set(1)
        self.processing = False
        self.update_rename_button_state()
        # 重命名完成后，按钮文字自动变为"撤销重命名"
        self._sync_rename_button()
        if errors:
            with open("rename_errors.log", "w", encoding="utf-8") as f:
                f.write("\n---\n".join(errors))
            self.set_status("重命名完成，部分失败详见 rename_errors.log")
        else:
            self.set_status("重命名完成")

    def undo_rename(self) -> None:
        """撤销最近一次重命名操作。"""
        if not self.rename_history:
            self.set_status("没有可撤销的操作。")
            return
        if self.processing or getattr(self, "_scanning", False):
            self.set_status("正在处理中，无法撤销。")
            return

        self.processing = True
        self.rename_button.configure(state="disabled")
        self.root.update_idletasks()

        success = failed = 0
        # 按逆序撤销（从最后一个重命名开始）
        for source_path, target_path, target_name in reversed(self.rename_history):
            if not os.path.exists(target_path):
                failed += 1
                continue
            try:
                os.rename(target_path, source_path)
                # 更新 records 中对应的记录
                for record in self.records:
                    if record.get("path") == target_path:
                        record["path"] = source_path
                        record["current_name"] = os.path.basename(source_path)
                        break
                success += 1
            except Exception:
                failed += 1

        self.rename_history.clear()
        self.processing = False

        # 刷新表格显示
        self.refresh_ui_after_undo()
        self._sync_rename_button()
        self.update_rename_button_state()

        if failed:
            self.set_status("撤销完成，部分文件已不存在")
        else:
            self.set_status("撤销完成")

    def refresh_ui_after_undo(self) -> None:
        """撤销后刷新表格。"""
        items = self.tree.get_children()
        for idx, record in enumerate(self.records):
            if idx < len(items):
                item = items[idx]
                self.tree.set(item, "new", record["current_name"])
                self.tree.set(item, "status", "已撤销")
                self.tree.item(item, tags=("idle",))
        self.update_stats()

    # ── 文件选择 ────────────────────────────────────────────────────────

    def load_from_argv(self) -> None:
        if len(sys.argv) <= 1:
            return
        incoming = sys.argv[1]
        if os.path.isdir(incoming):
            self.selected_paths = []
            self.path_var.set(incoming)
            self.root.after(150, self.scan_files)
        elif os.path.isfile(incoming) and incoming.lower().endswith(ALLOWED_EXTENSIONS):
            self.selected_paths = [incoming]
            self.path_var.set(os.path.basename(incoming))
            self.root.after(150, self.scan_files)

    # ── 云端 OCR 状态管理 ──────────────────────────────────────────────

    def _load_cloud_ocr_state(self) -> None:
        """加载云端 OCR 配置。"""
        try:
            from cloud_ocr import load_credentials
            creds = load_credentials()
            self._cloud_ocr_enabled = creds.get("enabled", False)
            self._cloud_secret_id = creds.get("secret_id", "")
            self._cloud_secret_key = creds.get("secret_key", "")
        except Exception:
            pass
        self._update_cloud_status_dot()

    def _update_cloud_status_dot(self) -> None:
        """更新云端 OCR 状态按钮外观与点击行为。"""
        if not hasattr(self, "_cloud_status_btn"):
            return
        has_creds = self._has_cloud_creds()
        if not has_creds:
            self._cloud_status_btn.configure(
                text="● 未配置",
                text_color="#FA5151",
                fg_color="#FCEBEB",
                hover_color="#FCEBEB",
                cursor="arrow",
            )
            return

        if self._cloud_ocr_enabled:
            self._cloud_status_btn.configure(
                text="● 已启用",
                text_color="#3FA9F5",
                fg_color="#E8F4FD",
                hover_color="#D6EAF8",
                cursor="hand2",
            )
        else:
            self._cloud_status_btn.configure(
                text="● 未启用",
                text_color="#8A8A8A",
                fg_color="#F2F3F5",
                hover_color="#E5E5E5",
                cursor="hand2",
            )

    def _has_cloud_creds(self) -> bool:
        return bool(self._cloud_secret_id and self._cloud_secret_key)

    def _open_cloud_ocr_settings(self) -> None:
        """弹出云端 OCR 设置窗口（开关置顶、单卡片、底部按钮固定）。"""
        from cloud_ocr import (
            save_credentials, clear_credentials,
            validate_credentials, get_usage_stats,
        )

        top = ctk.CTkToplevel(self.root)
        top.title("云端 OCR 识别设置")
        top.geometry("480x460")
        top.resizable(False, False)
        top.transient(self.root)
        top.grab_set()

        # 主容器：浅灰背景
        container = ctk.CTkFrame(top, fg_color="#F5F5F5")
        container.pack(fill="both", expand=True)

        # ── 底部按钮栏（先 pack，固定占用 60px）──
        footer = ctk.CTkFrame(container, fg_color="#FFFFFF", height=60, corner_radius=0)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        left_group = ctk.CTkFrame(footer, fg_color="transparent")
        left_group.pack(side="left", padx=16, pady=14)

        right_group = ctk.CTkFrame(footer, fg_color="transparent")
        right_group.pack(side="right", padx=16, pady=14)

        # ── 中间内容区（填充剩余空间）──
        body = ctk.CTkFrame(container, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=16, pady=12)

        # === 单一卡片：开关 + API 密钥 ===
        card = ctk.CTkFrame(body, fg_color="#FFFFFF", corner_radius=8)
        card.pack(fill="both", expand=True)

        # 启用开关（置顶，天蓝色，56x24）
        switch_row = ctk.CTkFrame(card, fg_color="transparent")
        switch_row.pack(fill="x", padx=16, pady=(16, 10))
        switch_row.grid_columnconfigure(0, weight=1)

        switch_left = ctk.CTkFrame(switch_row, fg_color="transparent")
        switch_left.grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(
            switch_left, text="启用云端识别",
            font=self.font_body, text_color="#191919",
        ).pack(anchor="w")
        ctk.CTkLabel(
            switch_left, text="自动识别图片发票和扫描件 PDF",
            font=self.font_hint, text_color="#5F5E5A",
        ).pack(anchor="w")

        enable_switch = ctk.CTkSwitch(
            switch_row, text="",
            switch_width=56, switch_height=24,
            progress_color="#3FA9F5",
            fg_color="#D9D9D9",
            button_color="#FFFFFF", button_hover_color="#F7F7F7",
        )
        enable_switch.grid(row=0, column=1, sticky="e")
        enable_switch.select() if self._cloud_ocr_enabled else enable_switch.deselect()

        # 细分割线
        sep = ctk.CTkFrame(card, height=1, fg_color="#E5E5E5")
        sep.pack(fill="x", padx=16, pady=(0, 12))

        # API 密钥标题行：左侧标题 + 右侧用量（两行）
        title_row = ctk.CTkFrame(card, fg_color="transparent")
        title_row.pack(fill="x", padx=16, pady=(0, 10))
        title_row.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            title_row, text="API 密钥",
            font=self.font_section, text_color="#191919",
        ).grid(row=0, column=0, sticky="w")

        usage = get_usage_stats()
        used = usage["used"]
        remaining = usage["remaining"]
        limit = usage["limit"]
        if remaining <= 0:
            usage_color = "#FA5151"
        elif remaining <= 100:
            usage_color = "#FA9D3B"
        else:
            usage_color = "#8A8A8A"
        usage_lbl = ctk.CTkLabel(
            title_row,
            text=f"本月已调用 {used} 次\n免费额度 {limit} 次/月",
            font=self.font_hint, text_color=usage_color, anchor="e", justify="right",
        )
        usage_lbl.grid(row=0, column=1, sticky="e")

        # SecretId
        ctk.CTkLabel(
            card, text="SecretId", anchor="w",
            font=self.font_body, text_color="#4C4C4C",
        ).pack(fill="x", padx=16, pady=(0, 4))

        secret_id_entry = ctk.CTkEntry(
            card, placeholder_text="AKIDxxxxxxxxxxxxxxxxxxxx",
            font=self.font_body, height=36, corner_radius=6,
            fg_color="#FFFFFF", border_color="#E5E5E5",
        )
        secret_id_entry.pack(fill="x", padx=16, pady=(0, 2))
        secret_id_entry.insert(0, self._cloud_secret_id)

        # SecretId 提示行 + 获取链接
        sid_hint_row = ctk.CTkFrame(card, fg_color="transparent")
        sid_hint_row.pack(fill="x", padx=16, pady=(0, 12))
        ctk.CTkLabel(
            sid_hint_row, text="从腾讯云控制台 > API 密钥管理获取",
            font=self.font_hint, text_color="#5F5E5A", anchor="w",
        ).pack(side="left")
        link_lbl = ctk.CTkLabel(
            sid_hint_row, text="点击获取 API 密钥",
            font=self.font_hint, text_color="#07C160", anchor="e", cursor="hand2",
        )
        link_lbl.pack(side="right")
        link_lbl.bind("<Button-1>", lambda e: os.startfile("https://console.cloud.tencent.com/cam/capi"))

        # SecretKey + 无缝衔接的文字显示/隐藏按钮
        ctk.CTkLabel(
            card, text="SecretKey", anchor="w",
            font=self.font_body, text_color="#4C4C4C",
        ).pack(fill="x", padx=16, pady=(0, 4))

        sk_row = ctk.CTkFrame(card, fg_color="transparent")
        sk_row.pack(fill="x", padx=16, pady=(0, 2))
        sk_row.grid_columnconfigure(0, weight=1)

        sk_container = ctk.CTkFrame(
            sk_row, height=38, corner_radius=6,
            fg_color="#FFFFFF", border_width=1, border_color="#E5E5E5",
        )
        sk_container.grid(row=0, column=0, sticky="ew")
        sk_container.grid_propagate(False)
        sk_container.grid_columnconfigure(0, weight=1)

        secret_key_entry = ctk.CTkEntry(
            sk_container, placeholder_text="xxxxxxxxxxxxxxxxxxxxxxxx",
            font=self.font_body, height=36, border_width=0, corner_radius=0,
            fg_color="#FFFFFF", show="*",
        )
        secret_key_entry.grid(row=0, column=0, sticky="ew", padx=(10, 0))
        secret_key_entry.insert(0, self._cloud_secret_key)

        def _toggle_key_visibility():
            if secret_key_entry.cget("show") == "*":
                secret_key_entry.configure(show="")
                toggle_btn.configure(text="隐藏")
            else:
                secret_key_entry.configure(show="*")
                toggle_btn.configure(text="显示")

        toggle_btn = ctk.CTkButton(
            sk_container, text="显示",
            width=50, height=36, corner_radius=6,
            fg_color="#FFFFFF", hover_color="#F2F3F5",
            text_color="#191919", font=self.font_hint,
            command=_toggle_key_visibility,
        )
        toggle_btn.grid(row=0, column=1, sticky="e", padx=(0, 1))

        ctk.CTkLabel(
            card, text="保存时自动加密混淆，防止明文泄露",
            font=self.font_hint, text_color="#5F5E5A", anchor="w",
        ).pack(fill="x", padx=16, pady=(0, 4))

        def _update_switch(*_):
            has_both = bool(secret_id_entry.get().strip() and secret_key_entry.get().strip())
            enable_switch.configure(state="normal" if has_both else "disabled")
            if not has_both:
                enable_switch.deselect()
        secret_id_entry.bind("<KeyRelease>", _update_switch)
        secret_key_entry.bind("<KeyRelease>", _update_switch)
        _update_switch()

        def _do_clear():
            secret_id_entry.delete(0, "end")
            secret_key_entry.delete(0, "end")
            enable_switch.deselect()
            _update_switch()
            clear_credentials()
            self._cloud_ocr_enabled = False
            self._cloud_secret_id = ""
            self._cloud_secret_key = ""
            self._update_cloud_status_dot()
            self.update_rename_button_state()

        def _do_verify():
            sid = secret_id_entry.get().strip()
            skey = secret_key_entry.get().strip()
            if not sid or not skey:
                messagebox.showinfo("验证", "请先输入 SecretId 和 SecretKey")
                return
            valid, msg = validate_credentials(sid, skey)
            if valid:
                messagebox.showinfo("验证结果", msg)
            else:
                messagebox.showwarning("验证结果", msg)

        def _do_save():
            sid = secret_id_entry.get().strip()
            skey = secret_key_entry.get().strip()
            if not sid or not skey:
                messagebox.showinfo("保存", "请先输入 SecretId 和 SecretKey")
                return
            enabled = bool(enable_switch.get())
            save_credentials(sid, skey, enabled=enabled)
            self._cloud_ocr_enabled = enabled
            self._cloud_secret_id = sid
            self._cloud_secret_key = skey
            self._update_cloud_status_dot()
            self.update_rename_button_state()
            top.destroy()

        ctk.CTkButton(
            left_group, text="清除密钥",
            width=80, height=36, corner_radius=6,
            fg_color="#F2F3F5", hover_color="#EDEDED",
            text_color="#FA5151", font=self.font_button,
            command=_do_clear,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            left_group, text="验证密钥",
            width=72, height=36, corner_radius=6,
            fg_color="#FFFFFF", border_width=1, border_color="#E5E5E5",
            text_color="#191919", font=self.font_body,
            command=_do_verify,
        ).pack(side="left")

        ctk.CTkButton(
            right_group, text="取消",
            width=60, height=36, corner_radius=6,
            fg_color="#FFFFFF", border_width=1, border_color="#E5E5E5",
            text_color="#191919", font=self.font_body,
            command=top.destroy,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            right_group, text="保存",
            width=60, height=36, corner_radius=6,
            fg_color="#3FA9F5", hover_color="#3498DB",
            text_color="#FFFFFF", font=self.font_button,
            command=_do_save,
        ).pack(side="left")

    def _popup_source_menu(self) -> None:
        """点击「选择来源」按钮后弹出自定义风格化下拉菜单。

        用透明色 trick 实现圆角：Toplevel 背景设为透明色，
        -transparentcolor 让该色变透明，内部 CTkFrame 的圆角区域外
        显示 Toplevel 背景（透明），从而露出圆角效果。
        """
        if hasattr(self, "_source_menu") and self._source_menu is not None:
            return

        btn = self._source_btn
        x = btn.winfo_rootx()
        y = btn.winfo_rooty() + btn.winfo_height() + 4
        menu_w = 130
        _TRANSPARENT = "#01a2b1"  # 不可能出现在 UI 中的颜色

        top = tk.Toplevel(self.root)
        top.overrideredirect(True)
        top.configure(bg=_TRANSPARENT)
        top.attributes("-topmost", True)
        top.attributes("-transparentcolor", _TRANSPARENT)

        # 带圆角 + 边框的容器
        container = ctk.CTkFrame(
            top, fg_color="#FFFFFF", corner_radius=8,
            border_width=1, border_color="#D0D0D0",
        )
        container.pack(fill="both", expand=True, padx=2, pady=2)

        _text_font = ctk.CTkFont(family="Microsoft YaHei UI", size=12)
        _text_font_bold = ctk.CTkFont(family="Microsoft YaHei UI", size=12, weight="bold")
        items = [
            ("选择文件夹", self.choose_folder, True),   # 主操作，加粗
            ("选择文件", self.choose_files, False),
            None,  # 分隔线
            ("清空", self._clear_source, False),
        ]

        for item in items:
            if item is None:
                sep = tk.Frame(container, height=1, bg="#E5E5E5")
                sep.pack(fill="x", padx=12, pady=2)
                continue
            text, callback, bold = item

            row = ctk.CTkFrame(container, fg_color="transparent", corner_radius=6)
            row.pack(fill="x", padx=5, pady=1)

            text_lbl = ctk.CTkLabel(
                row, text=text,
                font=_text_font_bold if bold else _text_font,
                text_color="#191919",
                anchor="w", height=30,
            )
            text_lbl.pack(side="left", fill="x", expand=True, padx=(12, 8))

            # 整行可点击 + hover 效果
            def _on_enter(e, r=row):
                r.configure(fg_color="#F0F0F0")
            def _on_leave(e, r=row):
                r.configure(fg_color="transparent")
            def _on_click(e, c=callback):
                self._menu_select(c)

            for w in (row, text_lbl):
                w.bind("<Enter>", _on_enter)
                w.bind("<Leave>", _on_leave)
                w.bind("<Button-1>", _on_click)

        # 精确高度：根据内容自适应，+4 保证底部边框不被裁切
        top.update_idletasks()
        real_h = top.winfo_reqheight() + 4
        top.geometry(f"{menu_w}x{real_h}+{x}+{y}")
        self._source_menu = top

        # 点击外部 / 失焦关闭
        top.bind("<FocusOut>", lambda e: self._close_source_menu())
        top.after(10, lambda: top.focus_force())

    def _menu_select(self, callback) -> None:
        """选中菜单项后先关闭菜单再执行回调。"""
        self._close_source_menu()
        callback()

    def _close_source_menu(self) -> None:
        """关闭下拉菜单。"""
        if hasattr(self, "_source_menu") and self._source_menu is not None:
            self._source_menu.destroy()
            self._source_menu = None

    def _clear_source(self) -> None:
        self.selected_paths = []
        self.path_var.set("")
        self.records.clear()
        self.rename_history.clear()
        self.clear_table()
        self._reset_sort_state()
        self.update_stats()
        self.update_rename_button_state()
        self._sync_rename_button()
        self.set_status("已清空")
        self.progress_bar.set(0)

    def choose_folder(self) -> None:
        folder = filedialog.askdirectory(title="请选择包含发票 PDF 或图片的文件夹")
        if folder:
            self.selected_paths = []
            self.path_var.set(folder)
            self.scan_files()

    def choose_files(self) -> None:
        files = filedialog.askopenfilenames(
            title="选择发票文件",
            filetypes=[
                ("发票文件", "*.pdf *.jpg *.jpeg *.png *.bmp *.tiff"),
                ("PDF 文件", "*.pdf"),
                ("图片文件", "*.jpg *.jpeg *.png *.bmp *.tiff"),
            ],
        )
        if files:
            self.selected_paths = list(files)
            ext_counts = {}
            for f in files:
                ext = os.path.splitext(f)[1].lower()
                ext_counts[ext] = ext_counts.get(ext, 0) + 1
            summary = ", ".join(f"{c} {e}" for e, c in ext_counts.items())
            self.path_var.set(f"已选择 {len(files)} 个文件 ({summary})")
            self.scan_files()

    def resolve_input_paths(self) -> list[str]:
        if self.selected_paths:
            return [
                p for p in self.selected_paths
                if os.path.isfile(p) and p.lower().endswith(ALLOWED_EXTENSIONS)
            ]
        source = self.path_var.get().strip()
        if source and os.path.isdir(source):
            return [
                os.path.join(source, n)
                for n in sorted(os.listdir(source))
                if n.lower().endswith(ALLOWED_EXTENSIONS)
            ]
        if source and os.path.isfile(source) and source.lower().endswith(ALLOWED_EXTENSIONS):
            return [source]
        return []

    # ── 命名相关 ────────────────────────────────────────────────────────

    def current_enabled_map(self) -> dict:
        return {k: bool(v.get()) for k, v in self.template_vars.items()}

    def compose_name(self, fields: dict, ext: str = ".pdf") -> str:
        return build_name(
            fields, self.current_enabled_map(), self.field_order,
            self.custom_value_var.get(), ext=ext,
        )

    def is_unrecognized_record(self, record: dict) -> bool:
        if record.get("not_invoice"):
            return True  # 非发票文件，重命名时跳过
        if record.get("error"):
            return True
        fields = record["fields"]
        return not any(fields.get(k) for k in ["date", "number", "buyer", "seller", "amount", "type"])

    def apply_name_to_record(self, record: dict) -> str:
        if self.is_unrecognized_record(record):
            record["new_name"] = record["current_name"]
            return record["new_name"]
        ext = os.path.splitext(record["path"])[1].lower() or ".pdf"
        desired = self.compose_name(record["fields"], ext=ext)
        directory = os.path.dirname(record["path"])
        record["new_name"] = ensure_unique_name(
            directory, desired, current_name=record["current_name"]
        )
        return record["new_name"]

    def on_template_change(self) -> None:
        if self.records:
            self.refresh_record_names()

    def refresh_record_names(self) -> None:
        items = self.tree.get_children()
        for idx, record in enumerate(self.records):
            self.apply_name_to_record(record)
            if idx < len(items):
                self.tree.set(items[idx], "new", record["new_name"])

    # ── 扫描文件 ────────────────────────────────────────────────────────

    def clear_table(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

    def classify_record(self, record: dict) -> tuple[str, str]:
        if record.get("not_invoice"):
            return "not_invoice", "非发票"
        if record["error"]:
            return "error", record["error"]
        fields = record["fields"]
        has_amount = bool(fields.get("amount"))
        has_subject = bool(fields.get("seller") or fields.get("buyer"))
        has_any = any(fields.get(k) for k in ["date", "number", "buyer", "seller", "amount", "type"])
        if has_amount and has_subject:
            return "success", "完整识别"
        if has_any:
            return "partial", "部分识别"
        return "idle", "未识别"

    def update_stats(self) -> None:
        total = len(self.records)
        complete = partial = failed = not_invoice = 0
        for record in list(self.records):  # 快照迭代，避免并发修改
            tag, _ = self.classify_record(record)
            if tag == "success":
                complete += 1
            elif tag == "partial":
                partial += 1
            elif tag == "not_invoice":
                not_invoice += 1
            else:
                failed += 1
        self.stats_vars["total"].set(str(total))
        self.stats_vars["complete"].set(str(complete))
        self.stats_vars["partial"].set(str(partial))
        self.stats_vars["failed"].set(str(failed))
        self.stats_vars["not_invoice"].set(str(not_invoice))

    def scan_files(self) -> None:
        if getattr(self, "_scanning", False):
            return
        paths = self.resolve_input_paths()
        if not paths:
            return

        self._scanning = True
        self.rename_history.clear()  # 新扫描时清除旧的撤销记录
        self.update_rename_button_state()
        self.set_status("正在扫描...")
        self.progress_bar.set(0)
        self.clear_table()
        self.records.clear()
        self._reset_sort_state()  # 新扫描清除排序
        self._scan_total = len(paths)
        self._scan_done_count = 0

        # 主线程捕获 UI 状态快照（避免工作线程访问 Tkinter 变量）
        enabled_map = self.current_enabled_map()
        field_order = list(self.field_order)
        custom_value = self.custom_value_var.get()

        def process_one(file_path: str) -> dict:
            """处理单个文件：判定 → 解析 → 生成命名。返回 record。"""
            ext = os.path.splitext(file_path)[1].lower()
            if ext in IMAGE_EXTENSIONS:
                if self._has_cloud_creds() and self._cloud_ocr_enabled:
                    parsed = parse_image_cloud(
                        file_path, self._cloud_secret_id, self._cloud_secret_key,
                    )
                else:
                    parsed = {"fields": {}, "error": "图片文件（未配置云端OCR）", "not_invoice": True}
            else:
                parsed = parse_invoice(file_path)
                # PDF 解析失败（无文字/扫描件/非发票）→ 兜底走云端识别
                if (parsed.get("not_invoice")
                        and self._has_cloud_creds() and self._cloud_ocr_enabled):
                    cloud_parsed = parse_image_cloud(
                        file_path, self._cloud_secret_id, self._cloud_secret_key,
                    )
                    if not cloud_parsed.get("not_invoice") and not cloud_parsed.get("error"):
                        parsed = cloud_parsed
            record = {
                "path": file_path,
                "source_name": os.path.basename(file_path),
                "current_name": os.path.basename(file_path),
                "new_name": "",
                "fields": parsed["fields"],
                "error": parsed["error"],
                "not_invoice": parsed["not_invoice"],
            }
            if self.is_unrecognized_record(record):
                record["new_name"] = record["current_name"]
            else:
                desired = build_name(record["fields"], enabled_map, field_order, custom_value,
                                      ext=ext if ext in IMAGE_EXTENSIONS else ".pdf")
                directory = os.path.dirname(file_path)
                record["new_name"] = ensure_unique_name(
                    directory, desired, current_name=record["current_name"],
                )
            return record

        def worker() -> None:
            for idx, file_path in enumerate(paths, start=1):
                record = process_one(file_path)
                self.records.append(record)
                self._scan_done_count = idx
                # 逐行插入表格（主线程执行，线程安全）
                self.root.after(0, lambda i=idx, r=record: self.insert_record_row(i, r))
            self.root.after(0, self._scan_finished)

        threading.Thread(target=worker, daemon=True).start()
        # 定时器更新进度条和统计卡片
        self._poll_scan_progress()

    def _poll_scan_progress(self) -> None:
        """主线程定时器：更新进度条和统计卡片。"""
        if not getattr(self, "_scanning", False):
            return  # 扫描已结束，不再覆盖状态
        done = self._scan_done_count
        total = self._scan_total
        if total > 0:
            self.progress_bar.set(done / total)
            self.set_status(f"扫描中... ({done}/{total})")
            self.update_stats()  # 实时刷新统计卡片
            self.root.after(200, self._poll_scan_progress)

    def _scan_finished(self) -> None:
        """扫描全部完成：收尾工作（行已逐条插入）。"""
        self._scanning = False
        self.progress_bar.set(1)
        self.auto_fit_columns()
        self.update_stats()
        self.set_status(f"扫描完成")
        self.update_rename_button_state()
        self._sync_rename_button()

    def insert_record_row(self, idx: int, record: dict) -> None:
        tag, status_text = self.classify_record(record)
        fields = record["fields"]
        subject = fields.get("seller") or fields.get("buyer") or "-"
        amount = f"¥{fields.get('amount')}" if fields.get("amount") else "-"
        self.tree.insert(
            "", tk.END,
            values=(idx, record["source_name"], record["new_name"],
                    fields.get("type", "") or "-", subject, amount, status_text),
            tags=(tag,),
        )

    # ── 表头排序 ────────────────────────────────────────────────────────

    def _reset_sort_state(self) -> None:
        """重置排序状态，恢复表头原始文字。"""
        self._sort_col = None
        self._sort_reverse = False
        base_headings = {
            "idx": ("#", tk.CENTER), "org": ("原文件名", tk.W),
            "new": ("新文件名", tk.W), "type": ("发票类型", tk.W),
            "seller": ("销售方/行程信息", tk.W), "amount": ("金额", tk.CENTER),
            "status": ("状态", tk.CENTER),
        }
        for k, (text, anchor) in base_headings.items():
            self.tree.heading(k, text=text, anchor=anchor)

    # 各列排序键提取函数（返回用于比较的值）
    _SORT_KEYS = {
        "idx":     lambda r: r[0],  # 在 _sort_by_column 中特殊处理
        "org":     lambda r: r[1]["source_name"],
        "new":     lambda r: r[1]["new_name"],
        "type":    lambda r: r[1]["fields"].get("type", "") or "",
        "seller":  lambda r: r[1]["fields"].get("seller") or r[1]["fields"].get("buyer") or "",
        "amount":  lambda r: float(r[1]["fields"].get("amount", 0) or 0),
        "status":  lambda r: r[2],  # status_text，在调用时传入
    }

    def _sort_by_column(self, col: str) -> None:
        """点击表头排序：首次点击升序，再次点击切换方向。"""
        # 扫描中不允许排序
        if getattr(self, "_scanning", False):
            return
        if not self.records:
            return

        # 切换排序方向
        if self._sort_col == col:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_col = col
            self._sort_reverse = False

        # 重置所有表头文字（去掉旧的 ▲▼）
        base_headings = {
            "idx": "#", "org": "原文件名", "new": "新文件名",
            "type": "发票类型", "seller": "销售方/行程信息",
            "amount": "金额", "status": "状态",
        }
        arrow = " ▼" if self._sort_reverse else " ▲"
        for k, text in base_headings.items():
            display = text + (arrow if k == self._sort_col else "")
            anchor = tk.CENTER if k in ("idx", "amount", "status") else tk.W
            self.tree.heading(k, text=display, anchor=anchor)

        # 构建排序数据：(原始序号, record, status_text)
        classified = []
        for i, record in enumerate(self.records):
            tag, status_text = self.classify_record(record)
            classified.append((i, record, status_text))

        key_fn = self._SORT_KEYS.get(col)
        if key_fn is None:
            return
        if col == "status":
            classified.sort(key=lambda x: x[2], reverse=self._sort_reverse)
        elif col == "idx":
            classified.sort(key=lambda x: x[0], reverse=self._sort_reverse)
        else:
            classified.sort(key=key_fn, reverse=self._sort_reverse)

        # 同步更新 self.records 顺序
        self.records = [item[1] for item in classified]

        # 重建表格（idx 重新编号）
        self.clear_table()
        for new_idx, record in enumerate(self.records, start=1):
            self.insert_record_row(new_idx, record)
        self.auto_fit_columns()

    def auto_fit_columns(self) -> None:
        """根据实际内容自动调整每列宽度，取表头和所有单元格的最大文字宽度 + 内边距。"""
        col_ids = ["idx", "org", "new", "type", "seller", "amount", "status"]
        # 用 Treeview 实际渲染字体来测量
        try:
            font_obj = tkfont.Font(self.tree, self.tree.cget("font"))
        except Exception:
            font_obj = tkfont.Font(family="Microsoft YaHei UI", size=10)

        # 基础表头文字（不含排序箭头），用于列宽计算
        _base_headings = {
            "idx": "#", "org": "原文件名", "new": "新文件名",
            "type": "发票类型", "seller": "销售方/行程信息",
            "amount": "金额", "status": "状态",
        }

        for col_id in col_ids:
            # 表头文字宽度（用基础文字，不含排序箭头）
            heading = _base_headings.get(col_id, "")
            max_w = font_obj.measure(heading)

            # 遍历所有行的该列内容
            for item in self.tree.get_children(""):
                val = str(self.tree.set(item, col_id))
                w = font_obj.measure(val)
                if w > max_w:
                    max_w = w

            # 内边距 28px（含列边距和排序箭头余量），上限 500 防止过宽
            new_width = min(max_w + 28, 500)
            # 保留下限 40
            new_width = max(new_width, 40)
            self.tree.column(col_id, width=new_width, minwidth=40, stretch=False)

    # ── Excel 导出 ─────────────────────────────────────────────────────

    def export_excel(self) -> None:
        export_invoice_excel(self.records, parent_window=self.root)


# ── 程序入口 ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    setup_windows_app_id()
    root = ctk.CTk()
    app = App(root)
    root.mainloop()
