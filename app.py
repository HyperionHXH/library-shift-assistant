# -*- coding: utf-8 -*-
"""
图书馆助理排班与微信群自动替班助手 (app.py)
桌面端图形界面主程序 (基于 Tkinter / ttk 高分屏适配)
"""

from __future__ import annotations
import ctypes
import datetime
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict, List, Optional, Tuple

from schedule_manager import (
    ScheduleManager,
    ShiftSlot,
    STANDARD_TIME_SLOTS,
    WEEKDAYS,
)
from message_parser import ShiftMessageParser
from wechat_monitor import WeChatMonitor, list_wechat_windows, WeChatWindowInfo

# 开启 Windows 高分屏 DPI 清晰渲染
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def get_app_dir() -> str:
    """获取程序运行根目录（兼容 PyInstaller 打包后的 exe 同级目录）"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


class LibraryShiftApp(tk.Tk):
    """图书馆助理自动替班与排班管理桌面客户端"""

    def __init__(self, default_excel: str = "示例排班表.xlsx"):
        super().__init__()
        self.title("📚 图书馆助理排班与微信群自动替班助手")
        scr_w = self.winfo_screenwidth()
        scr_h = self.winfo_screenheight()
        win_w = min(1320, max(960, scr_w - 60))
        win_h = min(840, max(640, scr_h - 80))
        self.geometry(f"{win_w}x{win_h}")
        self.minsize(920, 620)
        self.configure(bg="#F3F4F6")

        # 确定默认排班表路径：
        # 1) 优先加载同目录下非示例的真实排班表（含“排班”字样且非“示例排班表.xlsx”，方便放在 exe 同目录自动生效）
        # 2) 若没有外部真实表，加载同目录下的 示例排班表.xlsx，或 exe 内置打包的默认示例排班表
        # 3) 若仍未找到，再兜底扫描同目录/工作目录下的其他 .xlsx
        base_dir = get_app_dir()
        bundle_dir = getattr(sys, "_MEIPASS", None)
        excel_path = ""

        candidates: List[str] = []
        if os.path.isdir(base_dir):
            base_xlsx = sorted(
                os.path.join(base_dir, f)
                for f in os.listdir(base_dir)
                if f.lower().endswith(".xlsx") and not f.startswith("~$") and "最新排班" not in f
            )
            candidates.extend(
                p for p in base_xlsx
                if "排班" in os.path.basename(p) and os.path.basename(p) != default_excel
            )
            local_default = os.path.join(base_dir, default_excel)
            if os.path.exists(local_default):
                candidates.append(local_default)
        else:
            base_xlsx = []

        if bundle_dir and os.path.isdir(bundle_dir):
            bundled_default = os.path.join(bundle_dir, default_excel)
            if os.path.exists(bundled_default):
                candidates.append(bundled_default)

        candidates.extend(p for p in base_xlsx if p not in candidates)
        if not getattr(sys, "frozen", False) and os.path.isdir(os.getcwd()):
            for f in os.listdir(os.getcwd()):
                if f.lower().endswith(".xlsx") and not f.startswith("~$") and "最新排班" not in f:
                    p = os.path.join(os.getcwd(), f)
                    if p not in candidates:
                        candidates.append(p)

        excel_path = candidates[0] if candidates else os.path.join(base_dir, default_excel)

        self.schedule_mgr = ScheduleManager(
            excel_path=excel_path,
            state_path=os.path.join(base_dir, "schedule_state.json"),
        )
        self.parser = ShiftMessageParser(
            schedule_mgr=self.schedule_mgr,
            rules_path=os.path.join(base_dir, "parser_rules.json"),
        )
        self.monitor = WeChatMonitor(
            on_new_message=self._on_wechat_message_threadsafe,
            on_log=self._on_monitor_log_threadsafe,
            on_status_change=self._on_monitor_status_threadsafe,
        )

        self.wx_windows: List[WeChatWindowInfo] = []
        self.matrix_buttons: Dict[Tuple[str, str], tk.Label] = {}
        self._last_date_iso: str = datetime.date.today().isoformat()

        self._setup_styles()
        self._build_top_bar()
        self._build_main_notebook()
        self._build_status_bar()

        # 初始化界面数据
        self._refresh_all_views()
        self._refresh_wechat_window_list()
        self._tick_clock()

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------
    # UI 样式定义
    # ------------------------------------------------------------------
    def _setup_styles(self) -> None:
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure("TNotebook", background="#F3F4F6")
        style.configure("TNotebook.Tab", font=("Microsoft YaHei UI", 10, "bold"), padding=(14, 6))
        style.configure("Treeview", font=("Microsoft YaHei UI", 9), rowheight=26)
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 9, "bold"))
        style.configure("TLabelframe", background="#FFFFFF")
        style.configure("TLabelframe.Label", font=("Microsoft YaHei UI", 10, "bold"), background="#FFFFFF", foreground="#1F2937")

    # ------------------------------------------------------------------
    # 顶部控制栏（姓名输入、实时时钟、状态概览、一键保存与监控启停）
    # ------------------------------------------------------------------
    def _build_top_bar(self) -> None:
        top_frame = tk.Frame(self, bg="#1E293B", padx=14, pady=10)
        top_frame.pack(side=tk.TOP, fill=tk.X)

        # 左侧：标题与实时日期时钟
        left_box = tk.Frame(top_frame, bg="#1E293B")
        left_box.pack(side=tk.LEFT)

        tk.Label(
            left_box,
            text="📚 图书馆助理自动替班系统",
            font=("Microsoft YaHei UI", 13, "bold"),
            fg="#F8FAFC",
            bg="#1E293B",
        ).pack(anchor="w")

        self.clock_label = tk.Label(
            left_box,
            text="🕒 正在同步实时日期...",
            font=("Microsoft YaHei UI", 9),
            fg="#38BDF8",
            bg="#1E293B",
        )
        self.clock_label.pack(anchor="w", pady=(2, 0))

        # 中间：操作者姓名选择与读取
        mid_box = tk.Frame(top_frame, bg="#1E293B", padx=18)
        mid_box.pack(side=tk.LEFT)

        tk.Label(
            mid_box,
            text="👤 操作者姓名：",
            font=("Microsoft YaHei UI", 10, "bold"),
            fg="#E2E8F0",
            bg="#1E293B",
        ).grid(row=0, column=0, sticky="w")

        self.operator_var = tk.StringVar(value=self.schedule_mgr.operator_name)
        self.operator_combo = ttk.Combobox(
            mid_box,
            textvariable=self.operator_var,
            values=self.schedule_mgr.all_names,
            width=12,
            font=("Microsoft YaHei UI", 10),
        )
        self.operator_combo.grid(row=0, column=1, padx=4)
        self.operator_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_select_operator())
        self.operator_combo.bind("<Return>", lambda _e: self._on_select_operator())

        tk.Button(
            mid_box,
            text="🔍 读取我的排班",
            font=("Microsoft YaHei UI", 9, "bold"),
            bg="#2563EB",
            fg="#FFFFFF",
            activebackground="#1D4ED8",
            activeforeground="#FFFFFF",
            relief=tk.FLAT,
            padx=10,
            pady=3,
            cursor="hand2",
            command=self._on_select_operator,
        ).grid(row=0, column=2, padx=4)

        tk.Button(
            mid_box,
            text="💾 保存筛选配置",
            font=("Microsoft YaHei UI", 9, "bold"),
            bg="#059669",
            fg="#FFFFFF",
            activebackground="#047857",
            activeforeground="#FFFFFF",
            relief=tk.FLAT,
            padx=10,
            pady=3,
            cursor="hand2",
            command=self._on_save_config_clicked,
        ).grid(row=0, column=3, padx=4)

        self.summary_label = tk.Label(
            mid_box,
            text="请先选择或输入您的姓名以加载个人排班",
            font=("Microsoft YaHei UI", 9),
            fg="#FDE047",
            bg="#1E293B",
        )
        self.summary_label.grid(row=1, column=0, columnspan=4, sticky="w", pady=(4, 0))

        # 右侧：监控快捷开关与文件操作
        right_box = tk.Frame(top_frame, bg="#1E293B")
        right_box.pack(side=tk.RIGHT)

        self.top_monitor_btn = tk.Button(
            right_box,
            text="🚀 启动微信群监控",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg="#10B981",
            fg="#FFFFFF",
            activebackground="#059669",
            activeforeground="#FFFFFF",
            relief=tk.FLAT,
            padx=14,
            pady=6,
            cursor="hand2",
            command=self._toggle_monitor,
        )
        self.top_monitor_btn.pack(side=tk.RIGHT, padx=(8, 0))

        file_btn_box = tk.Frame(right_box, bg="#1E293B")
        file_btn_box.pack(side=tk.RIGHT, padx=6)

        tk.Button(
            file_btn_box,
            text="📂 加载排班表(.xlsx)",
            font=("Microsoft YaHei UI", 8),
            bg="#334155",
            fg="#E2E8F0",
            relief=tk.FLAT,
            padx=8,
            pady=2,
            cursor="hand2",
            command=self._on_choose_excel,
        ).pack(fill=tk.X, pady=(0, 3))

        tk.Button(
            file_btn_box,
            text="📤 导出最新排班(.xlsx)",
            font=("Microsoft YaHei UI", 8),
            bg="#334155",
            fg="#E2E8F0",
            relief=tk.FLAT,
            padx=8,
            pady=2,
            cursor="hand2",
            command=self._on_export_excel,
        ).pack(fill=tk.X)

    # ------------------------------------------------------------------
    # 主选项卡区域
    # ------------------------------------------------------------------
    def _build_main_notebook(self) -> None:
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=8)

        # Tab 1: 个人排班与禁接筛选矩阵
        self.tab_personal = tk.Frame(self.notebook, bg="#F3F4F6")
        self.notebook.add(self.tab_personal, text=" 📅 1. 个人排班与可接班筛选 ")
        self._build_tab_personal()

        # Tab 2: 全馆分岗位排班总表与单独设置
        self.tab_all_shifts = tk.Frame(self.notebook, bg="#F3F4F6")
        self.notebook.add(self.tab_all_shifts, text=" 🏢 2. 全馆排班总表与手动替班 ")
        self._build_tab_all_shifts()

        # Tab 3: 微信群监控与替班模拟测试台
        self.tab_monitor = tk.Frame(self.notebook, bg="#F3F4F6")
        self.notebook.add(self.tab_monitor, text=" 📡 3. 微信群实时监控与模拟测试台 ")
        self._build_tab_monitor()

        # Tab 4: 替班历史记录与语料规则配置
        self.tab_records = tk.Frame(self.notebook, bg="#F3F4F6")
        self.notebook.add(self.tab_records, text=" 📋 4. 替班记录(防搞混/撤销)与语料规则 ")
        self._build_tab_records()

    # ------------------------------------------------------------------
    # Tab 1: 个人排班与可接班筛选矩阵
    # ------------------------------------------------------------------
    def _build_tab_personal(self) -> None:
        paned = tk.PanedWindow(self.tab_personal, orient=tk.HORIZONTAL, bg="#E5E7EB", sashwidth=5)
        paned.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        # 左栏：本人当前排班表 + 岗位勾选
        left_panel = tk.Frame(paned, bg="#FFFFFF", padx=10, pady=10)
        paned.add(left_panel, width=410)

        tk.Label(
            left_panel,
            text="📌 我的当前排班表（含原固定班次与新接替班）",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg="#FFFFFF",
            fg="#111827",
        ).pack(anchor="w")

        tk.Label(
            left_panel,
            text="提示：已有班次的同时间段及重叠时间段会自动锁定排除，防止撞班。",
            font=("Microsoft YaHei UI", 8),
            bg="#FFFFFF",
            fg="#6B7280",
        ).pack(anchor="w", pady=(0, 6))

        cols = ("weekday", "date", "time", "post", "source")
        self.my_shifts_tree = ttk.Treeview(left_panel, columns=cols, show="headings", height=11)
        self.my_shifts_tree.heading("weekday", text="星期")
        self.my_shifts_tree.heading("date", text="日期")
        self.my_shifts_tree.heading("time", text="时间段")
        self.my_shifts_tree.heading("post", text="岗位")
        self.my_shifts_tree.heading("source", text="状态/来源")

        self.my_shifts_tree.column("weekday", width=48, anchor="center")
        self.my_shifts_tree.column("date", width=68, anchor="center")
        self.my_shifts_tree.column("time", width=95, anchor="center")
        self.my_shifts_tree.column("post", width=75, anchor="center")
        self.my_shifts_tree.column("source", width=95, anchor="center")
        self.my_shifts_tree.pack(fill=tk.BOTH, expand=True)

        # 岗位允许筛选区
        post_box = tk.LabelFrame(
            left_panel,
            text="🏢 允许接班的岗位（取消勾选则不接该岗位）",
            font=("Microsoft YaHei UI", 9, "bold"),
            bg="#FFFFFF",
            fg="#1F2937",
            padx=8,
            pady=6,
        )
        post_box.pack(fill=tk.X, pady=(10, 4))

        self.post_vars: Dict[str, tk.BooleanVar] = {}
        posts = ["服务台", "2楼书库", "3楼书库", "4楼书库", "D5图书室"]
        for i, p in enumerate(posts):
            var = tk.BooleanVar(value=(p in self.schedule_mgr.allowed_posts))
            self.post_vars[p] = var
            cb = tk.Checkbutton(
                post_box,
                text=p,
                variable=var,
                font=("Microsoft YaHei UI", 9),
                bg="#FFFFFF",
                activebackground="#FFFFFF",
                command=self._on_post_filter_changed,
            )
            cb.grid(row=i // 3, column=i % 3, sticky="w", padx=6, pady=2)

        # 快捷批量筛选按钮区
        batch_box = tk.LabelFrame(
            left_panel,
            text="⚡ 快捷批量设置不可接时段",
            font=("Microsoft YaHei UI", 9, "bold"),
            bg="#FFFFFF",
            fg="#1F2937",
            padx=8,
            pady=6,
        )
        batch_box.pack(fill=tk.X, pady=(6, 0))

        btn_cfg = {"font": ("Microsoft YaHei UI", 8), "relief": tk.GROOVE, "padx": 6, "pady": 3, "cursor": "hand2"}
        tk.Button(
            batch_box, text="🚫 禁接全部早八班(08:00起)", bg="#FEE2E2", fg="#991B1B",
            command=lambda: self._batch_block_time(lambda t, s, e, wd: s == 480), **btn_cfg
        ).grid(row=0, column=0, sticky="ew", padx=3, pady=2)

        tk.Button(
            batch_box, text="🚫 禁接全部20:00-23:00闭馆班", bg="#FEE2E2", fg="#991B1B",
            command=lambda: self._batch_block_time(lambda t, s, e, wd: t == "20:00-23:00"), **btn_cfg
        ).grid(row=0, column=1, sticky="ew", padx=3, pady=2)

        tk.Button(
            batch_box, text="🚫 禁接周末全部班次(周六/周日)", bg="#FEE2E2", fg="#991B1B",
            command=lambda: self._batch_block_time(lambda t, s, e, wd: wd in ("周六", "周日")), **btn_cfg
        ).grid(row=1, column=0, sticky="ew", padx=3, pady=2)

        tk.Button(
            batch_box, text="♻️ 清空全部手动禁接设置", bg="#E0F2FE", fg="#075985",
            command=self._clear_all_blocked, **btn_cfg
        ).grid(row=1, column=1, sticky="ew", padx=3, pady=2)
        batch_box.columnconfigure(0, weight=1)
        batch_box.columnconfigure(1, weight=1)

        tk.Button(
            left_panel,
            text="👁️ 查看筛选后剩余可接班次列表 (全部岗位详情)",
            font=("Microsoft YaHei UI", 9, "bold"),
            bg="#0284C7",
            fg="#FFFFFF",
            activebackground="#0369A1",
            activeforeground="#FFFFFF",
            relief=tk.FLAT,
            pady=5,
            cursor="hand2",
            command=self._jump_to_available_shifts,
        ).pack(fill=tk.X, pady=(8, 0))

        # 右栏：全周可视化时间段矩阵（点击切换可接/不可接）
        right_panel = tk.Frame(paned, bg="#FFFFFF", padx=10, pady=10)
        paned.add(right_panel, width=900)

        header_row = tk.Frame(right_panel, bg="#FFFFFF")
        header_row.pack(fill=tk.X, pady=(0, 4))

        tk.Label(
            header_row,
            text="🗓️ 全周排班与时间段筛选图（点击格子切换【🟩可接】/【🟥不能接】）",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg="#FFFFFF",
            fg="#111827",
        ).pack(side=tk.LEFT)

        self.auto_sync_dates_var = tk.BooleanVar(value=self.schedule_mgr.auto_sync_week_dates)
        tk.Checkbutton(
            header_row,
            text="🔄 按实时日历自动同步本周(周一~周日)日期",
            variable=self.auto_sync_dates_var,
            bg="#FFFFFF",
            font=("Microsoft YaHei UI", 8, "bold"),
            fg="#0369A1",
            command=self._on_toggle_sync_week_dates,
        ).pack(side=tk.RIGHT)

        legend_frame = tk.Frame(right_panel, bg="#FFFFFF")
        legend_frame.pack(fill=tk.X, pady=(0, 6))
        for bg_c, fg_c, txt in [
            ("#DBEAFE", "#1E40AF", "🟦 本人原班"),
            ("#EDE9FE", "#5B21B6", "🟪 替班新接(已成功接下)"),
            ("#FEF3C7", "#92400E", "🟧 时间冲突"),
            ("#FEE2E2", "#991B1B", "🟥 设为不能接(点击恢复)"),
            ("#DCFCE7", "#166534", "🟩 允许接班(点击禁接)"),
        ]:
            tk.Label(
                legend_frame, text=txt, font=("Microsoft YaHei UI", 8, "bold"),
                bg=bg_c, fg=fg_c, padx=6, pady=2
            ).pack(side=tk.LEFT, padx=(0, 6))

        # 矩阵滚动容器
        canvas_frame = tk.Frame(right_panel, bg="#FFFFFF")
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.matrix_canvas = tk.Canvas(canvas_frame, bg="#FFFFFF", highlightthickness=0)
        v_scroll = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.matrix_canvas.yview)
        self.matrix_inner = tk.Frame(self.matrix_canvas, bg="#E5E7EB")

        self.matrix_inner.bind(
            "<Configure>",
            lambda _e: self.matrix_canvas.configure(scrollregion=self.matrix_canvas.bbox("all")),
        )
        self.matrix_canvas.create_window((0, 0), window=self.matrix_inner, anchor="nw")
        self.matrix_canvas.configure(yscrollcommand=v_scroll.set)

        self.matrix_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        def _on_mousewheel(event):
            self.matrix_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.matrix_canvas.bind("<Enter>", lambda _e: self.matrix_canvas.bind_all("<MouseWheel>", _on_mousewheel))
        self.matrix_canvas.bind("<Leave>", lambda _e: self.matrix_canvas.unbind_all("<MouseWheel>"))

        self._build_matrix_grid()

    def _build_matrix_grid(self) -> None:
        """构建 周一~周日 × 标准时段 的交互式网格"""
        for w in self.matrix_inner.winfo_children():
            w.destroy()
        self.matrix_buttons.clear()

        # 表头：时段说明 + 周一~周日（点击星期列标题可整天切换禁接！）
        tk.Label(
            self.matrix_inner,
            text="时段 \\ 星期 (点击列头整天禁接)",
            font=("Microsoft YaHei UI", 9, "bold"),
            bg="#334155",
            fg="#FFFFFF",
            width=22,
            pady=6,
        ).grid(row=0, column=0, padx=1, pady=1, sticky="nsew")

        today_iso = datetime.date.today().isoformat()
        for c_idx, wd in enumerate(WEEKDAYS, start=1):
            d_str, iso_d = self.schedule_mgr.weekday_dates.get(wd, ("", ""))
            is_today = (iso_d == today_iso)
            hdr_text = f"{wd} ({d_str})\n{'[今天]' if is_today else '点击整列切换'}"
            hdr_bg = "#0284C7" if is_today else "#334155"
            lbl = tk.Label(
                self.matrix_inner,
                text=hdr_text,
                font=("Microsoft YaHei UI", 9, "bold"),
                bg=hdr_bg,
                fg="#FFFFFF",
                width=13,
                pady=4,
                cursor="hand2",
            )
            lbl.grid(row=0, column=c_idx, padx=1, pady=1, sticky="nsew")
            lbl.bind("<Button-1>", lambda _e, w=wd: self._toggle_entire_weekday(w))

        # 每一行对应一个标准时段
        for r_idx, (t_str, s_min, e_min, desc) in enumerate(STANDARD_TIME_SLOTS, start=1):
            row_hdr = tk.Label(
                self.matrix_inner,
                text=f"{t_str}\n{desc}",
                font=("Microsoft YaHei UI", 8, "bold"),
                bg="#F8FAFC",
                fg="#334155",
                padx=6,
                pady=4,
                cursor="hand2",
            )
            row_hdr.grid(row=r_idx, column=0, padx=1, pady=1, sticky="nsew")
            row_hdr.bind("<Button-1>", lambda _e, ts=t_str: self._toggle_entire_timeslot_row(ts))

            for c_idx, wd in enumerate(WEEKDAYS, start=1):
                # 检查该星期+时段在全馆是否有任何排班
                has_any_slot = any(
                    s.weekday == wd and s.time_str == t_str
                    for s in self.schedule_mgr.slots.values()
                )
                if not has_any_slot:
                    empty_lbl = tk.Label(
                        self.matrix_inner,
                        text="—",
                        font=("Microsoft YaHei UI", 8),
                        bg="#F3F4F6",
                        fg="#9CA3AF",
                        width=13,
                        pady=6,
                    )
                    empty_lbl.grid(row=r_idx, column=c_idx, padx=1, pady=1, sticky="nsew")
                    continue

                cell = tk.Label(
                    self.matrix_inner,
                    text="可接",
                    font=("Microsoft YaHei UI", 8, "bold"),
                    bg="#DCFCE7",
                    fg="#166534",
                    width=13,
                    pady=6,
                    cursor="hand2",
                )
                cell.grid(row=r_idx, column=c_idx, padx=1, pady=1, sticky="nsew")
                cell.bind("<Button-1>", lambda _e, w=wd, ts=t_str: self._on_matrix_cell_clicked(w, ts))
                self.matrix_buttons[(wd, t_str)] = cell

    # ------------------------------------------------------------------
    # Tab 2: 全馆分岗位排班总表与手动替班
    # ------------------------------------------------------------------
    def _build_tab_all_shifts(self) -> None:
        top_bar = tk.Frame(self.tab_all_shifts, bg="#FFFFFF", padx=12, pady=8)
        top_bar.pack(fill=tk.X, padx=6, pady=(6, 2))

        tk.Label(
            top_bar,
            text="筛选岗位：",
            font=("Microsoft YaHei UI", 9, "bold"),
            bg="#FFFFFF",
        ).pack(side=tk.LEFT)

        self.filter_post_var = tk.StringVar(value="全部岗位")
        for p in ["全部岗位", "服务台", "2楼书库", "3楼书库", "4楼书库", "D5图书室"]:
            rb = ttk.Radiobutton(
                top_bar,
                text=p,
                value=p,
                variable=self.filter_post_var,
                command=self._refresh_all_shifts_table,
            )
            rb.pack(side=tk.LEFT, padx=6)

        tk.Label(
            top_bar,
            text=" |  筛选星期：",
            font=("Microsoft YaHei UI", 9, "bold"),
            bg="#FFFFFF",
        ).pack(side=tk.LEFT, padx=(10, 2))

        self.filter_day_var = tk.StringVar(value="全部星期")
        day_combo = ttk.Combobox(
            top_bar,
            textvariable=self.filter_day_var,
            values=["全部星期"] + WEEKDAYS,
            width=10,
            state="readonly",
        )
        day_combo.pack(side=tk.LEFT, padx=4)
        day_combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_all_shifts_table())

        self.only_avail_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            top_bar,
            text="仅看【🟩可接班】",
            variable=self.only_avail_var,
            bg="#FFFFFF",
            font=("Microsoft YaHei UI", 9, "bold"),
            fg="#166534",
            command=self._refresh_all_shifts_table,
        ).pack(side=tk.LEFT, padx=(8, 4))

        # 右侧操作按钮
        tk.Button(
            top_bar,
            text="🤝 手动登记：由我替选中班次",
            font=("Microsoft YaHei UI", 9, "bold"),
            bg="#2563EB",
            fg="#FFFFFF",
            relief=tk.FLAT,
            padx=10,
            pady=3,
            cursor="hand2",
            command=self._on_manual_take_selected_slot,
        ).pack(side=tk.RIGHT, padx=4)

        tk.Button(
            top_bar,
            text="🚫 切换选中班次【单独禁接/解除】",
            font=("Microsoft YaHei UI", 9),
            bg="#EF4444",
            fg="#FFFFFF",
            relief=tk.FLAT,
            padx=10,
            pady=3,
            cursor="hand2",
            command=self._on_toggle_single_slot_block,
        ).pack(side=tk.RIGHT, padx=4)

        # 排班总表 Treeview
        table_frame = tk.Frame(self.tab_all_shifts, bg="#FFFFFF", padx=6, pady=4)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

        cols = ("slot_id", "post", "weekday", "date", "time", "current_holders", "original_holders", "my_status", "reason")
        self.all_shifts_tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=22)
        headers = [
            ("slot_id", "班次ID", 150),
            ("post", "岗位", 85),
            ("weekday", "星期", 55),
            ("date", "日期", 75),
            ("time", "时间段", 105),
            ("current_holders", "当前值班人(替班后)", 140),
            ("original_holders", "原固定排班人", 140),
            ("my_status", "对我的状态", 110),
            ("reason", "判定说明 / 备注", 320),
        ]
        for col_id, title, w in headers:
            self.all_shifts_tree.heading(col_id, text=title)
            self.all_shifts_tree.column(col_id, width=w, anchor="center" if col_id != "reason" else "w")

        self.all_shifts_tree.tag_configure("AVAILABLE", background="#F0FDF4", foreground="#166534")
        self.all_shifts_tree.tag_configure("OWNED", background="#EFF6FF", foreground="#1E40AF")
        self.all_shifts_tree.tag_configure("CONFLICT", background="#FEF3C7", foreground="#92400E")
        self.all_shifts_tree.tag_configure("BLOCKED", background="#FEF2F2", foreground="#991B1B")
        self.all_shifts_tree.tag_configure("EXPIRED", background="#F3F4F6", foreground="#6B7280")

        y_sb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.all_shifts_tree.yview)
        self.all_shifts_tree.configure(yscrollcommand=y_sb.set)
        self.all_shifts_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y_sb.pack(side=tk.RIGHT, fill=tk.Y)

    # ------------------------------------------------------------------
    # Tab 3: 微信群实时监控与模拟测试台
    # ------------------------------------------------------------------
    def _build_tab_monitor(self) -> None:
        paned = tk.PanedWindow(self.tab_monitor, orient=tk.HORIZONTAL, bg="#E5E7EB", sashwidth=5)
        paned.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        # 左栏：微信连接设置 + 模拟消息测试台
        left_col = tk.Frame(paned, bg="#FFFFFF", padx=10, pady=8)
        paned.add(left_col, width=480)

        # 1. 微信监控参数设置
        wx_box = tk.LabelFrame(
            left_col,
            text="⚙️ PC微信群监控与秒回设置",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg="#FFFFFF",
            padx=8,
            pady=6,
        )
        wx_box.pack(fill=tk.X, pady=(0, 8))

        tk.Label(wx_box, text="选择微信窗口：", bg="#FFFFFF", font=("Microsoft YaHei UI", 9)).grid(row=0, column=0, sticky="w", pady=2)
        self.wx_win_var = tk.StringVar(value="自动检测当前打开的微信窗口")
        self.wx_win_combo = ttk.Combobox(wx_box, textvariable=self.wx_win_var, width=22, state="readonly")
        self.wx_win_combo.grid(row=0, column=1, sticky="ew", padx=4, pady=2)

        tk.Button(
            wx_box, text="🔄 刷新窗口", font=("Microsoft YaHei UI", 8),
            command=self._refresh_wechat_window_list, relief=tk.GROOVE, cursor="hand2"
        ).grid(row=0, column=2, padx=2, pady=2)

        tk.Label(wx_box, text="群名关键词(可选)：", bg="#FFFFFF", font=("Microsoft YaHei UI", 9)).grid(row=1, column=0, sticky="w", pady=2)
        self.group_kw_var = tk.StringVar(value=self.schedule_mgr.target_group_name)
        tk.Entry(wx_box, textvariable=self.group_kw_var, font=("Microsoft YaHei UI", 9), width=20).grid(row=1, column=1, sticky="ew", padx=4, pady=2)

        tk.Button(
            wx_box, text="🔍 诊断测试读取", font=("Microsoft YaHei UI", 8),
            bg="#E0F2FE", fg="#0369A1", command=self._on_diagnose_wechat, relief=tk.GROOVE, cursor="hand2"
        ).grid(row=1, column=2, padx=2, pady=2)

        tk.Label(wx_box, text="自动抢答回复词：", bg="#FFFFFF", font=("Microsoft YaHei UI", 9, "bold")).grid(row=2, column=0, sticky="w", pady=2)
        self.reply_text_var = tk.StringVar(value=self.schedule_mgr.reply_text)
        tk.Entry(wx_box, textvariable=self.reply_text_var, font=("Microsoft YaHei UI", 10, "bold"), fg="#059669", width=20).grid(row=2, column=1, sticky="ew", padx=4, pady=2)

        tk.Label(wx_box, text="我的微信昵称(防自接)：", bg="#FFFFFF", font=("Microsoft YaHei UI", 9)).grid(row=3, column=0, sticky="w", pady=2)
        self.wx_aliases_var = tk.StringVar(value="，".join(self.parser.my_wechat_aliases))
        alias_entry = tk.Entry(wx_box, textvariable=self.wx_aliases_var, font=("Microsoft YaHei UI", 9), width=20)
        alias_entry.grid(row=3, column=1, sticky="ew", padx=4, pady=2)
        tk.Label(wx_box, text="逗号隔开", bg="#FFFFFF", font=("Microsoft YaHei UI", 8), fg="#6B7280").grid(row=3, column=2, sticky="w")

        cb_opts = {"bg": "#FFFFFF", "font": ("Microsoft YaHei UI", 8), "wraplength": 430, "justify": tk.LEFT, "anchor": "w"}

        self.auto_send_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            wx_box, text="开启自动向微信窗口发送回复（取消则为仅提醒/演练模式）",
            variable=self.auto_send_var, command=self._on_auto_send_toggled, **cb_opts
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=1)

        self.auto_update_var = tk.BooleanVar(value=self.schedule_mgr.auto_update_on_reply)
        tk.Checkbutton(
            wx_box, text="抢答后自动更新我的个人班次（防止撞班，可一键撤销）",
            variable=self.auto_update_var, **cb_opts
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=1)

        self.take_only_one_var = tk.BooleanVar(value=self.parser.take_only_one_in_multi)
        tk.Checkbutton(
            wx_box, text="多班消息每次仅自动接 1 个班（防止连抢多班被当脚本风控）",
            variable=self.take_only_one_var, command=self._on_parser_opt_toggled, **cb_opts
        ).grid(row=6, column=0, columnspan=3, sticky="w", pady=1)

        self.allow_partial_combo_var = tk.BooleanVar(value=self.parser.allow_partial_combo)
        tk.Checkbutton(
            wx_box, text="连班/多班部分可接时拆分接取（如“6-11”只接“6-8”）",
            variable=self.allow_partial_combo_var, command=self._on_parser_opt_toggled, **cb_opts
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=1)

        self.partial_with_time_var = tk.BooleanVar(value=self.parser.partial_reply_with_time)
        tk.Checkbutton(
            wx_box, text="多班/部分接班时回复带上12小时制简写时间（如“我可以 周日9-12”）",
            variable=self.partial_with_time_var, command=self._on_parser_opt_toggled, **cb_opts
        ).grid(row=8, column=0, columnspan=3, sticky="w", pady=1)

        wx_box.columnconfigure(1, weight=1)

        # 2. 替班消息仿真测试器
        sim_box = tk.LabelFrame(
            left_col,
            text="🧪 微信群替班消息仿真测试台（免开微信即可验证）",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg="#FFFFFF",
            padx=8,
            pady=6,
        )
        sim_box.pack(fill=tk.BOTH, expand=True)

        sender_row = tk.Frame(sim_box, bg="#FFFFFF")
        sender_row.grid(row=0, column=1, sticky="ew", padx=4, pady=2)
        tk.Label(sim_box, text="发信人昵称：", bg="#FFFFFF", font=("Microsoft YaHei UI", 9)).grid(row=0, column=0, sticky="w", pady=2)
        self.sim_sender_var = tk.StringVar(value="张三")
        tk.Entry(sender_row, textvariable=self.sim_sender_var, font=("Microsoft YaHei UI", 9), width=13).pack(side=tk.LEFT)
        self.sim_is_self_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            sender_row, text="模拟我自己发的消息(右气泡)", variable=self.sim_is_self_var,
            bg="#FFFFFF", font=("Microsoft YaHei UI", 8), fg="#DC2626"
        ).pack(side=tk.LEFT, padx=(6, 0))

        tk.Label(sim_box, text="预设语料：", bg="#FFFFFF", font=("Microsoft YaHei UI", 9)).grid(row=1, column=0, sticky="w", pady=2)
        sample_msgs = [
            "周一8-10二楼书库，周二2-3:30三楼书库，周三6-10二楼书库求替",
            "周三早9-10:30 四楼\n周四晚6-8 四楼\n周五早9-10:30 四楼\n求替",
            "两点到六点（1400-1530 1530 1800）服务台求替",
            "15：30-18：00二楼服务台（原王五班）求替",
            "今晚6-11服务台求替",
            "周二12：30-2：00服务台求替",
            "十点半到十二点半服务台求替",
            "今天下午 2-4 四楼求替",
            "今晚20-23四楼书库求替",
            "我可以",
        ]
        self.sim_preset_combo = ttk.Combobox(sim_box, values=[m.replace("\n", " / ") for m in sample_msgs], width=30)
        self._sim_raw_samples = sample_msgs
        self.sim_preset_combo.grid(row=1, column=1, sticky="ew", padx=4, pady=2)
        self.sim_preset_combo.bind("<<ComboboxSelected>>", self._on_sim_preset_selected)

        tk.Label(sim_box, text="消息内容：", bg="#FFFFFF", font=("Microsoft YaHei UI", 9)).grid(row=2, column=0, sticky="nw", pady=2)
        self.sim_msg_text = tk.Text(sim_box, height=3, font=("Microsoft YaHei UI", 9), wrap=tk.WORD)
        self.sim_msg_text.grid(row=2, column=1, sticky="ew", padx=4, pady=2)
        self.sim_msg_text.insert("1.0", sample_msgs[0])

        sim_btn_row = tk.Frame(sim_box, bg="#FFFFFF")
        sim_btn_row.grid(row=3, column=0, columnspan=2, sticky="ew", pady=4)

        tk.Button(
            sim_btn_row,
            text="🔍 测试解析判定（不改表）",
            font=("Microsoft YaHei UI", 9, "bold"),
            bg="#0284C7",
            fg="#FFFFFF",
            relief=tk.FLAT,
            padx=8,
            pady=4,
            cursor="hand2",
            command=lambda: self._run_simulation(apply_if_match=False),
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            sim_btn_row,
            text="⚡ 模拟抢答并更新排班",
            font=("Microsoft YaHei UI", 9, "bold"),
            bg="#059669",
            fg="#FFFFFF",
            relief=tk.FLAT,
            padx=8,
            pady=4,
            cursor="hand2",
            command=lambda: self._run_simulation(apply_if_match=True),
        ).pack(side=tk.LEFT)

        self.sim_result_label = tk.Label(
            sim_box,
            text="点击上方按钮查看解析详情与是否接班判定...",
            font=("Microsoft YaHei UI", 9),
            bg="#F8FAFC",
            fg="#1E293B",
            justify=tk.LEFT,
            anchor="nw",
            wraplength=430,
            padx=8,
            pady=6,
            relief=tk.SOLID,
            bd=1,
        )
        self.sim_result_label.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=4)
        sim_box.columnconfigure(1, weight=1)
        sim_box.rowconfigure(4, weight=1)

        # 右栏：实时监控日志
        right_col = tk.Frame(paned, bg="#FFFFFF", padx=12, pady=10)
        paned.add(right_col, width=760)

        log_hdr = tk.Frame(right_col, bg="#FFFFFF")
        log_hdr.pack(fill=tk.X, pady=(0, 6))

        tk.Label(
            log_hdr,
            text="📜 微信群监控与自动抢答实时日志",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg="#FFFFFF",
        ).pack(side=tk.LEFT)

        tk.Button(
            log_hdr,
            text="🗑️ 清空日志",
            font=("Microsoft YaHei UI", 8),
            command=lambda: self.log_text.delete("1.0", tk.END),
            relief=tk.GROOVE,
        ).pack(side=tk.RIGHT)

        self.log_text = tk.Text(
            right_col,
            font=("Consolas", 10),
            bg="#0F172A",
            fg="#E2E8F0",
            insertbackground="#FFFFFF",
            wrap=tk.WORD,
            padx=10,
            pady=8,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.tag_configure("TAKE", foreground="#4ADE80", font=("Consolas", 10, "bold"))
        self.log_text.tag_configure("SKIP", foreground="#FBBF24")
        self.log_text.tag_configure("INFO", foreground="#38BDF8")
        self.log_text.tag_configure("ERR", foreground="#F87171")

        self._append_log("💡 欢迎使用图书馆助理替班助手！请先在上方选择您的姓名并设置不可接时段，随后即可启动监控或在左侧模拟测试。", "INFO")

    # ------------------------------------------------------------------
    # Tab 4: 替班历史记录与语料规则配置
    # ------------------------------------------------------------------
    def _build_tab_records(self) -> None:
        paned = tk.PanedWindow(self.tab_records, orient=tk.VERTICAL, bg="#E5E7EB", sashwidth=5)
        paned.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        # 上半部分：替班历史记录与一键撤销
        top_pane = tk.Frame(paned, bg="#FFFFFF", padx=12, pady=10)
        paned.add(top_pane, height=340)

        hdr = tk.Frame(top_pane, bg="#FFFFFF")
        hdr.pack(fill=tk.X, pady=(0, 6))

        tk.Label(
            hdr,
            text="📋 替班更新历史记录（精准记录原班主与接班人，若群里被他人抢先可一键撤销恢复原状）",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg="#FFFFFF",
        ).pack(side=tk.LEFT)

        tk.Button(
            hdr,
            text="↩️ 撤销选中的替班记录（恢复原排班）",
            font=("Microsoft YaHei UI", 9, "bold"),
            bg="#DC2626",
            fg="#FFFFFF",
            relief=tk.FLAT,
            padx=12,
            pady=4,
            cursor="hand2",
            command=self._on_undo_selected_record,
        ).pack(side=tk.RIGHT)

        cols = ("record_id", "timestamp", "weekday", "date", "post", "time", "original_person", "operator_name", "status", "source_message")
        self.records_tree = ttk.Treeview(top_pane, columns=cols, show="headings", height=10)
        r_headers = [
            ("record_id", "记录编号", 145),
            ("timestamp", "替班时间", 135),
            ("weekday", "星期", 55),
            ("date", "日期", 70),
            ("post", "岗位", 80),
            ("time", "班次时段", 130),
            ("original_person", "被替班人(原班主)", 110),
            ("operator_name", "接班人(我)", 90),
            ("status", "状态", 75),
            ("source_message", "触发微信消息", 280),
        ]
        for cid, title, w in r_headers:
            self.records_tree.heading(cid, text=title)
            self.records_tree.column(cid, width=w, anchor="center" if cid != "source_message" else "w")
        self.records_tree.pack(fill=tk.BOTH, expand=True)

        # 下半部分：语料解析规则自定义
        bot_pane = tk.LabelFrame(
            paned,
            text="🧠 替班语料判定词库配置（可根据微信群聊天习惯随时增删关键词，用逗号分隔）",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg="#FFFFFF",
            padx=12,
            pady=10,
        )
        paned.add(bot_pane, height=280)

        tk.Label(bot_pane, text="求替班触发词：", bg="#FFFFFF", font=("Microsoft YaHei UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        self.req_kw_var = tk.StringVar(value="，".join(self.parser.request_keywords))
        tk.Entry(bot_pane, textvariable=self.req_kw_var, font=("Microsoft YaHei UI", 9), width=90).grid(row=0, column=1, sticky="ew", padx=8, pady=5)

        tk.Label(bot_pane, text="忽略/防误触词：", bg="#FFFFFF", font=("Microsoft YaHei UI", 9, "bold")).grid(row=1, column=0, sticky="w", pady=5)
        self.ign_kw_var = tk.StringVar(value="，".join(self.parser.ignore_keywords))
        tk.Entry(bot_pane, textvariable=self.ign_kw_var, font=("Microsoft YaHei UI", 9), width=90).grid(row=1, column=1, sticky="ew", padx=8, pady=5)

        self.ignore_swap_var = tk.BooleanVar(value=self.parser.ignore_swap_only)
        tk.Checkbutton(
            bot_pane,
            text="忽略纯“换班/互换/调班”请求（只接纯替班，防止把自己的班换出去）",
            variable=self.ignore_swap_var,
            bg="#FFFFFF",
            font=("Microsoft YaHei UI", 9),
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=3)

        tk.Label(bot_pane, text="AI语义兜底(可选)：", bg="#FFFFFF", font=("Microsoft YaHei UI", 9, "bold"), fg="#475569").grid(row=3, column=0, sticky="w", pady=5)
        llm_row = tk.Frame(bot_pane, bg="#FFFFFF")
        llm_row.grid(row=3, column=1, sticky="ew", padx=8, pady=5)
        tk.Label(llm_row, text="API URL:", bg="#FFFFFF", font=("Microsoft YaHei UI", 8)).pack(side=tk.LEFT)
        self.llm_api_url_var = tk.StringVar(value=self.parser.llm_api_url)
        tk.Entry(llm_row, textvariable=self.llm_api_url_var, font=("Microsoft YaHei UI", 9), width=32).pack(side=tk.LEFT, padx=(2, 8), fill=tk.X, expand=True)
        tk.Label(llm_row, text="Model:", bg="#FFFFFF", font=("Microsoft YaHei UI", 8)).pack(side=tk.LEFT)
        self.llm_model_var = tk.StringVar(value=self.parser.llm_model)
        tk.Entry(llm_row, textvariable=self.llm_model_var, font=("Microsoft YaHei UI", 9), width=16).pack(side=tk.LEFT, padx=(2, 8))
        tk.Label(llm_row, text="API Key:", bg="#FFFFFF", font=("Microsoft YaHei UI", 8)).pack(side=tk.LEFT)
        self.llm_api_key_var = tk.StringVar(value=self.parser.llm_api_key)
        tk.Entry(llm_row, textvariable=self.llm_api_key_var, font=("Microsoft YaHei UI", 9), width=18, show="*").pack(side=tk.LEFT, padx=2)

        tk.Button(
            bot_pane,
            text="💾 保存语料规则与 AI 兜底配置",
            font=("Microsoft YaHei UI", 9, "bold"),
            bg="#2563EB",
            fg="#FFFFFF",
            relief=tk.FLAT,
            padx=14,
            pady=4,
            cursor="hand2",
            command=self._on_save_rules_clicked,
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=6)
        bot_pane.columnconfigure(1, weight=1)

    # ------------------------------------------------------------------
    # 底部状态栏
    # ------------------------------------------------------------------
    def _build_status_bar(self) -> None:
        bar = tk.Frame(self, bg="#E2E8F0", padx=12, pady=4)
        bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.status_bar_label = tk.Label(
            bar,
            text=f"📄 当前排班文件: {os.path.basename(self.schedule_mgr.excel_path)} ({self.schedule_mgr.week_title})",
            font=("Microsoft YaHei UI", 8),
            bg="#E2E8F0",
            fg="#334155",
        )
        self.status_bar_label.pack(side=tk.LEFT)

        self.monitor_state_label = tk.Label(
            bar,
            text="⚪ 监控状态：未启动",
            font=("Microsoft YaHei UI", 8, "bold"),
            bg="#E2E8F0",
            fg="#475569",
        )
        self.monitor_state_label.pack(side=tk.RIGHT)

    # ------------------------------------------------------------------
    # 实时时钟与同步逻辑
    # ------------------------------------------------------------------
    def _tick_clock(self) -> None:
        now = datetime.datetime.now()
        cur_iso = now.date().isoformat()
        if cur_iso != self._last_date_iso:
            self._last_date_iso = cur_iso
            if self.schedule_mgr.auto_sync_week_dates:
                self.schedule_mgr.sync_dates_to_current_week(use_realtime=True)
            self._build_matrix_grid()
            self._refresh_all_views()
            self.status_bar_label.configure(
                text=f"📄 当前排班文件: {os.path.basename(self.schedule_mgr.excel_path)} ({self.schedule_mgr.week_title})"
            )

        wd = WEEKDAYS[now.weekday()]
        clock_str = (
            f"🕒 实时同步：{now.strftime('%Y-%m-%d %H:%M:%S')} 【今天是 {wd}】 "
            f"| 当前排班周：{self.schedule_mgr.week_title}"
        )
        self.clock_label.configure(text=clock_str)
        self.after(1000, self._tick_clock)

    def _on_toggle_sync_week_dates(self) -> None:
        use_rt = self.auto_sync_dates_var.get()
        self.schedule_mgr.sync_dates_to_current_week(use_realtime=use_rt)
        self.schedule_mgr.save_state()
        self._build_matrix_grid()
        self._refresh_all_views()
        self.status_bar_label.configure(
            text=f"📄 当前排班文件: {os.path.basename(self.schedule_mgr.excel_path)} ({self.schedule_mgr.week_title})"
        )

    # ------------------------------------------------------------------
    # 数据刷新与交互事件处理
    # ------------------------------------------------------------------
    def _on_select_operator(self) -> None:
        name = self.operator_var.get().strip()
        self.schedule_mgr.operator_name = name
        self.schedule_mgr.save_state()
        self._refresh_all_views()
        if name:
            my_slots = self.schedule_mgr.get_person_current_slots(name)
            self._append_log(f"👤 已切换操作者为【{name}】，当前共有 {len(my_slots)} 个班次。", "INFO")

    def _save_config(self, show_popup: bool = True) -> None:
        self.schedule_mgr.operator_name = self.operator_var.get().strip()
        self.schedule_mgr.reply_text = self.reply_text_var.get().strip() or "我可以"
        self.schedule_mgr.target_group_name = self.group_kw_var.get().strip()
        self.schedule_mgr.auto_update_on_reply = self.auto_update_var.get()
        self.schedule_mgr.ignore_expired_shifts = False
        self.schedule_mgr.save_state()

        self.parser.allow_partial_combo = self.allow_partial_combo_var.get()
        self.parser.partial_reply_with_time = self.partial_with_time_var.get()
        self.parser.take_only_one_in_multi = self.take_only_one_var.get()
        self.parser.my_wechat_aliases = [
            x.strip() for x in self.wx_aliases_var.get().replace(",", "，").split("，") if x.strip()
        ]
        self.parser.save_rules()

        self._refresh_all_views()
        if show_popup:
            messagebox.showinfo("保存成功", "个人排班筛选条件（已有班次排除、不可接时段、允许岗位、微信昵称防自接、自动回复设置）已保存！")

    def _on_parser_opt_toggled(self) -> None:
        self.parser.allow_partial_combo = self.allow_partial_combo_var.get()
        self.parser.partial_reply_with_time = self.partial_with_time_var.get()
        self.parser.take_only_one_in_multi = self.take_only_one_var.get()
        self.parser.save_rules()

    def _on_save_config_clicked(self) -> None:
        self._save_config(show_popup=True)

    def _jump_to_available_shifts(self) -> None:
        """保存当前筛选并跳转到 Tab 2 仅显示剩余可接班次列表"""
        self._save_config(show_popup=False)
        self.filter_post_var.set("全部岗位")
        self.filter_day_var.set("全部星期")
        self.only_avail_var.set(True)
        self._refresh_all_shifts_table()
        self.notebook.select(self.tab_all_shifts)

    def _on_auto_send_toggled(self) -> None:
        self.monitor.auto_send_enabled = self.auto_send_var.get()

    def _on_post_filter_changed(self) -> None:
        allowed = {p for p, var in self.post_vars.items() if var.get()}
        self.schedule_mgr.allowed_posts = allowed
        self.schedule_mgr.save_state()
        self._refresh_all_views()

    def _on_matrix_cell_clicked(self, weekday: str, time_str: str) -> None:
        """点击全周矩阵中的某个时段单元格，切换【可接】与【不可接】"""
        op = self.schedule_mgr.operator_name.strip()
        _, s_min, e_min = next(
            ((t, sm, em) for t, sm, em, _ in STANDARD_TIME_SLOTS if t == time_str),
            (time_str, 0, 0),
        )
        if op:
            conflict = self.schedule_mgr.is_time_window_conflicting_with_existing(op, weekday, s_min, e_min)
            if conflict:
                is_sub = not conflict.was_originally_held_by(op)
                if is_sub and conflict.time_str == time_str:
                    messagebox.showinfo(
                        "已替班接下",
                        f"您已替班接下【{weekday} {conflict.time_str} {conflict.post}】！\n"
                        f"如需退回/撤销，请前往第 4 选项卡【替班记录】选中后点击撤销。",
                    )
                else:
                    messagebox.showwarning(
                        "已有班次冲突",
                        f"您在【{weekday} {conflict.time_str}】已有【{conflict.post}】的班次！\n"
                        f"系统已自动锁定排除该时间段及重叠时间段，防止撞班。",
                    )
                return

        key = f"{weekday}|{time_str}"
        if key in self.schedule_mgr.blocked_time_keys:
            self.schedule_mgr.blocked_time_keys.remove(key)
        else:
            self.schedule_mgr.blocked_time_keys.add(key)

        self.schedule_mgr.save_state()
        self._refresh_all_views()

    def _toggle_entire_weekday(self, weekday: str) -> None:
        """点击星期列标题，整列切换该天所有时段的禁接状态"""
        day_keys = [f"{weekday}|{t}" for t, _, _, _ in STANDARD_TIME_SLOTS]
        any_unblocked = any(k not in self.schedule_mgr.blocked_time_keys for k in day_keys)
        if any_unblocked:
            for k in day_keys:
                self.schedule_mgr.blocked_time_keys.add(k)
        else:
            for k in day_keys:
                self.schedule_mgr.blocked_time_keys.discard(k)
        self.schedule_mgr.save_state()
        self._refresh_all_views()

    def _toggle_entire_timeslot_row(self, time_str: str) -> None:
        """点击左侧时段行标题，整行切换周一至周日该时段的禁接状态"""
        row_keys = [f"{wd}|{time_str}" for wd in WEEKDAYS]
        any_unblocked = any(k not in self.schedule_mgr.blocked_time_keys for k in row_keys)
        if any_unblocked:
            for k in row_keys:
                self.schedule_mgr.blocked_time_keys.add(k)
        else:
            for k in row_keys:
                self.schedule_mgr.blocked_time_keys.discard(k)
        self.schedule_mgr.save_state()
        self._refresh_all_views()

    def _batch_block_time(self, predicate) -> None:
        for t_str, s_min, e_min, _ in STANDARD_TIME_SLOTS:
            for wd in WEEKDAYS:
                if predicate(t_str, s_min, e_min, wd):
                    self.schedule_mgr.blocked_time_keys.add(f"{wd}|{t_str}")
        self.schedule_mgr.save_state()
        self._refresh_all_views()

    def _clear_all_blocked(self) -> None:
        self.schedule_mgr.blocked_time_keys.clear()
        self.schedule_mgr.blocked_slot_ids.clear()
        self.schedule_mgr.allowed_posts = {"服务台", "2楼书库", "3楼书库", "4楼书库", "D5图书室"}
        for p, var in self.post_vars.items():
            var.set(True)
        self.schedule_mgr.save_state()
        self._refresh_all_views()

    def _refresh_all_views(self) -> None:
        """刷新个人班表、全周矩阵、全馆排班表、替班历史与顶部统计"""
        op = self.schedule_mgr.operator_name.strip()

        # 1. 刷新个人班次列表（原班次蓝色、替班新接紫色高亮）
        self.my_shifts_tree.tag_configure("NEW_SUB", background="#EDE9FE", foreground="#5B21B6")
        self.my_shifts_tree.tag_configure("ORIG", background="#FFFFFF", foreground="#111827")
        for item in self.my_shifts_tree.get_children():
            self.my_shifts_tree.delete(item)

        my_slots = self.schedule_mgr.get_person_current_slots(op)
        new_sub_cnt = 0
        for s in my_slots:
            is_orig = s.was_originally_held_by(op)
            if not is_orig:
                new_sub_cnt += 1
            src = "原固定班次" if is_orig else "🟪 替班新接"
            self.my_shifts_tree.insert(
                "", tk.END, values=(s.weekday, s.date_str, s.time_str, s.post, src),
                tags=("ORIG" if is_orig else "NEW_SUB",),
            )

        # 2. 刷新全周时段矩阵颜色与文字（清晰区分：原班 🟦 / 新接替班 🟪 / 撞班冲突 🟧 / 禁接 🟥 / 可接 🟩）
        for (wd, t_str), lbl in self.matrix_buttons.items():
            _, s_min, e_min = next(
                ((t, sm, em) for t, sm, em, _ in STANDARD_TIME_SLOTS if t == t_str),
                (t_str, 0, 0),
            )
            owned_here = [
                s for s in my_slots
                if s.weekday == wd and s.time_str == t_str
            ]
            if owned_here:
                sub_here = [s for s in owned_here if not s.was_originally_held_by(op)]
                orig_here = [s for s in owned_here if s.was_originally_held_by(op)]
                if sub_here:
                    posts_str = "/".join(s.post for s in sub_here)
                    lbl.configure(text=f"🟪 替接:{posts_str}", bg="#EDE9FE", fg="#5B21B6")
                else:
                    posts_str = "/".join(s.post for s in orig_here)
                    lbl.configure(text=f"🟦 原班:{posts_str}", bg="#DBEAFE", fg="#1E40AF")
                continue

            conflict_slot = self.schedule_mgr.is_time_window_conflicting_with_existing(op, wd, s_min, e_min) if op else None
            if conflict_slot:
                lbl.configure(text=f"🟧 冲突({conflict_slot.time_str})", bg="#FEF3C7", fg="#92400E")
                continue

            if self.schedule_mgr.is_time_window_blocked_by_user(wd, s_min, e_min, t_str):
                lbl.configure(text="🟥 不能接", bg="#FEE2E2", fg="#991B1B")
            else:
                # 统计该 (星期, 时段) 下经过岗位与单班筛选后实际允许的岗位
                cell_slots = [
                    s for s in self.schedule_mgr.slots.values()
                    if s.weekday == wd and s.time_str == t_str
                ]
                avail_posts = [
                    s.post for s in cell_slots
                    if self.schedule_mgr.evaluate_slot_status(s, op, check_expired=False)[0] == "AVAILABLE"
                ]
                if not avail_posts:
                    lbl.configure(text="🟥 岗位/已禁", bg="#FEE2E2", fg="#991B1B")
                elif len(avail_posts) == 1:
                    lbl.configure(text=f"🟩 {avail_posts[0]}", bg="#DCFCE7", fg="#166534")
                else:
                    lbl.configure(text=f"🟩 可接({len(avail_posts)}岗)", bg="#DCFCE7", fg="#166534")

        # 3. 刷新全馆排班总表
        self._refresh_all_shifts_table()

        # 4. 刷新替班历史记录表
        self._refresh_records_table()

        # 5. 更新顶部统计摘要
        avail_cnt = 0
        blocked_cnt = 0
        for s in self.schedule_mgr.slots.values():
            st, _ = self.schedule_mgr.evaluate_slot_status(s, op, check_expired=False)
            if st == "AVAILABLE":
                avail_cnt += 1
            elif st in ("BLOCKED_TIME", "BLOCKED_SLOT", "BLOCKED_POST"):
                blocked_cnt += 1

        if op:
            orig_cnt = len(my_slots) - new_sub_cnt
            sub_txt = f" | 🟪 新接:{new_sub_cnt}" if new_sub_cnt else ""
            self.summary_label.configure(
                text=f"操作者:【{op}】| 🟦 原班:{orig_cnt}{sub_txt} | 🟥 禁接:{blocked_cnt} | 🟩 可接:{avail_cnt}"
            )
        else:
            self.summary_label.configure(
                text=f"⚠️ 尚未选择操作者姓名（全馆共 {len(self.schedule_mgr.slots)} 个班次，请在上方下拉框选择或输入您的名字）"
            )

    def _refresh_all_shifts_table(self) -> None:
        for item in self.all_shifts_tree.get_children():
            self.all_shifts_tree.delete(item)

        post_filter = self.filter_post_var.get()
        day_filter = self.filter_day_var.get()
        only_avail = self.only_avail_var.get()
        now = datetime.datetime.now()

        status_label_map = {
            "AVAILABLE": ("🟩 可接班", "AVAILABLE"),
            "OWNED": ("🟦 本人班次", "OWNED"),
            "CONFLICT": ("🟧 时间冲突", "CONFLICT"),
            "BLOCKED_POST": ("🟥 岗位已禁", "BLOCKED"),
            "BLOCKED_TIME": ("🟥 时段已禁", "BLOCKED"),
            "BLOCKED_SLOT": ("🟥 单独禁接", "BLOCKED"),
            "EXPIRED": ("⏰ 已过期", "EXPIRED"),
        }

        for sid in self.schedule_mgr.slots_order:
            s = self.schedule_mgr.slots[sid]
            if post_filter != "全部岗位" and s.post != post_filter:
                continue
            if day_filter != "全部星期" and s.weekday != day_filter:
                continue

            st_code, reason = self.schedule_mgr.evaluate_slot_status(
                s, check_expired=False, now=now
            )
            if only_avail and st_code != "AVAILABLE":
                continue
            st_text, tag = status_label_map.get(st_code, (st_code, "AVAILABLE"))
            if s.extra_info:
                reason = f"{reason} | {s.extra_info}"

            self.all_shifts_tree.insert(
                "",
                tk.END,
                iid=sid,
                values=(
                    sid,
                    s.post,
                    s.weekday,
                    s.date_str,
                    s.time_str,
                    "、".join(s.current_holders),
                    "、".join(s.original_holders),
                    st_text,
                    reason,
                ),
                tags=(tag,),
            )

    def _refresh_records_table(self) -> None:
        for item in self.records_tree.get_children():
            self.records_tree.delete(item)
        for rec in reversed(self.schedule_mgr.records):
            self.records_tree.insert(
                "",
                tk.END,
                iid=rec.record_id,
                values=(
                    rec.record_id,
                    rec.timestamp,
                    rec.weekday,
                    rec.date_str,
                    rec.post,
                    rec.time_str,
                    rec.original_person,
                    rec.operator_name,
                    rec.status,
                    rec.source_message,
                ),
            )

    # ------------------------------------------------------------------
    # 全馆表手动操作与替班记录撤销
    # ------------------------------------------------------------------
    def _on_toggle_single_slot_block(self) -> None:
        sel = self.all_shifts_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在表格中选中一个或多个班次！")
            return
        for sid in sel:
            if sid in self.schedule_mgr.blocked_slot_ids:
                self.schedule_mgr.blocked_slot_ids.remove(sid)
            else:
                self.schedule_mgr.blocked_slot_ids.add(sid)
        self.schedule_mgr.save_state()
        self._refresh_all_views()

    def _on_manual_take_selected_slot(self) -> None:
        op = self.schedule_mgr.operator_name.strip()
        if not op:
            messagebox.showwarning("未设置姓名", "请先在顶部输入或选择您的操作者姓名！")
            return
        sel = self.all_shifts_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在排班总表中选中您要接替的班次！")
            return

        slots = [self.schedule_mgr.slots[sid] for sid in sel if sid in self.schedule_mgr.slots]
        # 若包含双人岗，询问替哪一位
        requester = ""
        if len(slots) == 1 and len(slots[0].current_holders) > 1:
            holders = slots[0].current_holders
            ans = messagebox.askyesno(
                "双人班次确认",
                f"该班次当前有两位值班人：【{holders[0]}】和【{holders[1]}】。\n"
                f"点击【是】替【{holders[0]}】的班；点击【否】替【{holders[1]}】的班（绝不覆盖另一人的班次）。",
            )
            requester = holders[0] if ans else holders[1]

        ok, msg, _ = self.schedule_mgr.apply_substitution(
            slots_to_take=slots,
            operator_name=op,
            requester_name=requester,
            source_message="[手动登记替班]",
            auto_replied=False,
        )
        if ok:
            self._append_log(f"🤝 {msg}", "TAKE")
            self._refresh_all_views()
            messagebox.showinfo("替班成功", msg)
        else:
            messagebox.showwarning("无法替班", msg)

    def _on_undo_selected_record(self) -> None:
        sel = self.records_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在替班历史记录表中选中一条要撤销的记录！")
            return
        rec_id = sel[0]
        ok, msg = self.schedule_mgr.undo_substitution(rec_id)
        if ok:
            self._append_log(f"↩️ {msg}", "INFO")
            self._refresh_all_views()
            messagebox.showinfo("已撤销替班", msg)
        else:
            messagebox.showwarning("撤销失败", msg)

    def _on_save_rules_clicked(self) -> None:
        req_kws = [x.strip() for x in self.req_kw_var.get().replace(",", "，").split("，") if x.strip()]
        ign_kws = [x.strip() for x in self.ign_kw_var.get().replace(",", "，").split("，") if x.strip()]
        self.parser.request_keywords = req_kws
        self.parser.ignore_keywords = ign_kws
        self.parser.ignore_swap_only = self.ignore_swap_var.get()
        self.parser.llm_api_url = self.llm_api_url_var.get().strip()
        self.parser.llm_model = self.llm_model_var.get().strip()
        self.parser.llm_api_key = self.llm_api_key_var.get().strip()
        self.parser.save_rules()
        messagebox.showinfo("保存成功", "替班语料触发词、防误触词与 AI 兜底配置已更新并保存！")

    # ------------------------------------------------------------------
    # Excel 导入与导出
    # ------------------------------------------------------------------
    def _on_choose_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="选择每周排班表 Excel 文件",
            filetypes=[("Excel 排班表", "*.xlsx")],
        )
        if not path:
            return
        old_path = self.schedule_mgr.excel_path
        self.schedule_mgr.load_excel(path)
        if self.auto_sync_dates_var.get():
            self.schedule_mgr.sync_dates_to_current_week(use_realtime=True)
        if os.path.abspath(path) != os.path.abspath(old_path) and self.schedule_mgr.records:
            if messagebox.askyesno(
                "切换新排班表",
                "检测到您加载了新的排班表文件，是否清空上一份排班表的替班记录？\n"
                "（您的操作者姓名及每周常规不可接时段设置将会保留）",
            ):
                self.schedule_mgr.reset_week_substitutions()
        self.schedule_mgr.save_state()
        self.operator_combo.configure(values=self.schedule_mgr.all_names)
        self._build_matrix_grid()
        self._refresh_all_views()
        self.status_bar_label.configure(
            text=f"📄 当前排班文件: {os.path.basename(self.schedule_mgr.excel_path)} ({self.schedule_mgr.week_title})"
        )
        self._append_log(f"📂 已重新加载排班表：{os.path.basename(path)} ({self.schedule_mgr.week_title})", "INFO")

    def _on_export_excel(self) -> None:
        base, ext = os.path.splitext(self.schedule_mgr.excel_path)
        default_out = f"{base}_最新排班{ext}"
        path = filedialog.asksaveasfilename(
            title="导出更新后的排班表 Excel",
            initialfile=os.path.basename(default_out),
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        ok, msg = self.schedule_mgr.export_updated_excel(path)
        if ok:
            self._append_log(f"📤 已导出最新排班表至：{msg}", "INFO")
            messagebox.showinfo("导出成功", f"最新排班表已导出至：\n{msg}")
        else:
            messagebox.showerror("导出失败", msg)

    # ------------------------------------------------------------------
    # 微信窗口检测、模拟测试与实时监控回调
    # ------------------------------------------------------------------
    def _refresh_wechat_window_list(self) -> None:
        self.wx_windows = list_wechat_windows()
        labels = ["自动检测当前打开的微信窗口"] + [w.display_label() for w in self.wx_windows]
        self.wx_win_combo.configure(values=labels)
        if len(labels) > 1:
            self.wx_win_combo.current(1)
        else:
            self.wx_win_combo.current(0)

    def _get_selected_hwnd(self) -> Optional[int]:
        idx = self.wx_win_combo.current()
        if idx > 0 and idx - 1 < len(self.wx_windows):
            return self.wx_windows[idx - 1].hwnd
        return None

    def _on_diagnose_wechat(self) -> None:
        self._refresh_wechat_window_list()
        hwnd = self._get_selected_hwnd()
        kw = self.group_kw_var.get().strip()
        report = self.monitor.inspect_wechat_window(target_hwnd=hwnd, keyword=kw)
        self._append_log("🔍 [微信窗口诊断报告]\n" + report, "INFO")
        messagebox.showinfo("微信窗口诊断结果", report)

    def _on_sim_preset_selected(self, _event) -> None:
        idx = self.sim_preset_combo.current()
        if 0 <= idx < len(self._sim_raw_samples):
            val = self._sim_raw_samples[idx]
        else:
            val = self.sim_preset_combo.get()
        self.sim_msg_text.delete("1.0", tk.END)
        self.sim_msg_text.insert("1.0", val)
        # 根据示例语料匹配模拟发信人
        if "周一8-10二楼书库" in val:
            self.sim_sender_var.set("张三")
        elif "周三早9-10:30" in val:
            self.sim_sender_var.set("李四")
        elif "两点到六点" in val:
            self.sim_sender_var.set("王五")
        elif "原王五班" in val:
            self.sim_sender_var.set("赵六")
        elif "今晚6-11" in val:
            self.sim_sender_var.set("钱七")
        elif "12：30-2：00" in val:
            self.sim_sender_var.set("赵六")
        elif "十点半到十二点半" in val:
            self.sim_sender_var.set("孙八")
        elif "今天下午 2-4" in val:
            self.sim_sender_var.set("周九")
        elif "20-23四楼" in val:
            self.sim_sender_var.set("吴十")

    def _run_simulation(self, apply_if_match: bool = False) -> None:
        op = self.operator_var.get().strip()
        if op != self.schedule_mgr.operator_name:
            self._on_select_operator()
        if not self.schedule_mgr.operator_name:
            messagebox.showwarning("请先选择姓名", "请先在界面顶部选择或输入您的操作者姓名！")
            return

        self.schedule_mgr.ignore_expired_shifts = False
        self.parser.allow_partial_combo = self.allow_partial_combo_var.get()
        self.parser.partial_reply_with_time = self.partial_with_time_var.get()
        self.parser.take_only_one_in_multi = self.take_only_one_var.get()
        self.parser.my_wechat_aliases = [
            x.strip() for x in self.wx_aliases_var.get().replace(",", "，").split("，") if x.strip()
        ]

        sender = self.sim_sender_var.get().strip()
        msg = self.sim_msg_text.get("1.0", tk.END).strip()
        is_self = self.sim_is_self_var.get()
        reply_word = self.reply_text_var.get().strip() or "我可以"

        res = self.parser.parse_and_evaluate(
            msg,
            sender_name=sender,
            is_self_hint=is_self,
            base_reply_word=reply_word,
        )
        lines = [
            f"【原始消息】{'[🙋本人右气泡] ' if is_self else ''}[{sender}] {msg.replace(chr(10), ' / ')}",
            f"【替班意图】{'是 (求替班请求)' if res.is_shift_request else '否 (' + res.reject_reason + ')'}",
        ]
        if res.is_shift_request:
            lines.append(f"【拆出班次】共 {len(res.items)} 个班次项（✅ 可接 {len(res.accepted_items)} 个，❌ 跳过 {len(res.rejected_items)} 个）")
            lines.append(f"【判定结论】{'✅ 立即接班！建议回复：「' + res.suggested_reply + '」' if res.should_take else '❌ 不接此班'}")
            lines.append(f"【详细解析】{res.decision_reason}")

        if res.should_take and apply_if_match:
            lines.append(f"【动作执行】🚀 自动发送回复：「{res.suggested_reply}」")
            applied_msgs = []
            for item in res.accepted_items:
                ok, sub_msg, _ = self.schedule_mgr.apply_substitution(
                    slots_to_take=item.matched_slots,
                    operator_name=self.schedule_mgr.operator_name,
                    requester_name=item.detected_requester,
                    source_message=f"[{sender}] {msg}",
                    auto_replied=True,
                )
                if ok:
                    applied_msgs.append(sub_msg)
            if applied_msgs:
                lines.append("【排班更新】" + "；".join(applied_msgs))
                self._append_log(f"🧪 [模拟抢答成功] 已回复「{res.suggested_reply}」| " + "；".join(applied_msgs), "TAKE")
            self._refresh_all_views()
        else:
            tag = "TAKE" if res.should_take else "SKIP"
            self._append_log(f"🧪 [模拟测试] [{sender}] {msg.replace(chr(10), ' / ')} -> {res.decision_reason or res.reject_reason}", tag)

        self.sim_result_label.configure(
            text="\n".join(lines),
            bg="#F0FDF4" if res.should_take else "#FEF2F2",
            fg="#166534" if res.should_take else "#991B1B",
        )

    def _toggle_monitor(self) -> None:
        if self.monitor.is_running:
            self.monitor.stop()
            return

        op = self.operator_var.get().strip()
        if op != self.schedule_mgr.operator_name:
            self._on_select_operator()
        if not self.schedule_mgr.operator_name:
            messagebox.showwarning("请先选择姓名", "启动监控前，请先在顶部输入或选择您的【操作者姓名】并确认不可接班次！")
            return

        self._save_config(show_popup=False)
        self._refresh_wechat_window_list()
        hwnd = self._get_selected_hwnd()
        kw = self.group_kw_var.get().strip()

        ok, msg = self.monitor.start(
            target_hwnd=hwnd,
            target_group_keyword=kw,
            poll_interval_sec=0.25,
            auto_send_enabled=self.auto_send_var.get(),
        )
        if not ok:
            messagebox.showwarning("启动失败", msg)
        else:
            self.notebook.select(self.tab_monitor)

    def _on_wechat_message_threadsafe(self, sender: str, content: str, is_self: bool) -> None:
        self.after(0, lambda: self._handle_incoming_wechat_message(sender, content, is_self))

    def _on_monitor_log_threadsafe(self, text: str) -> None:
        self.after(0, lambda: self._append_log(text, "INFO"))

    def _on_monitor_status_threadsafe(self, is_running: bool, status_text: str) -> None:
        def _update():
            if is_running:
                self.top_monitor_btn.configure(text="⏹ 停止微信群监控", bg="#DC2626", activebackground="#B91C1C")
                self.monitor_state_label.configure(text=f"🟢 监控状态：{status_text}", fg="#059669")
            else:
                self.top_monitor_btn.configure(text="🚀 启动微信群监控", bg="#10B981", activebackground="#059669")
                self.monitor_state_label.configure(text=f"⚪ 监控状态：{status_text}", fg="#475569")
        self.after(0, _update)

    def _handle_incoming_wechat_message(self, sender: str, content: str, is_self: bool) -> None:
        """处理微信群/聊天窗新消息：防本人自接 -> 多班分段解析 -> 极速回复 -> 精准更新个人排班"""
        self.schedule_mgr.ignore_expired_shifts = False
        self.parser.allow_partial_combo = self.allow_partial_combo_var.get()
        self.parser.partial_reply_with_time = self.partial_with_time_var.get()
        self.parser.take_only_one_in_multi = self.take_only_one_var.get()
        self.parser.my_wechat_aliases = [
            x.strip() for x in self.wx_aliases_var.get().replace(",", "，").split("，") if x.strip()
        ]

        reply_word = self.reply_text_var.get().strip() or "我可以"
        res = self.parser.parse_and_evaluate(
            content,
            sender_name=sender,
            is_self_hint=is_self,
            base_reply_word=reply_word,
        )
        sdr_disp = f"[{sender}] " if sender else ""
        preview = content.replace("\n", " / ")

        if is_self or res.reject_reason.startswith("本人发送"):
            self._append_log(f"🙋 忽略自己发送的消息 {sdr_disp}{preview}", "INFO")
            return

        if not res.is_shift_request:
            self._append_log(f"💬 新消息 {sdr_disp}{preview} -> 忽略 ({res.reject_reason})", "INFO")
            return

        if not res.should_take:
            self._append_log(f"⏭️ 替班请求 {sdr_disp}{preview} -> 不接 ({res.decision_reason})", "SKIP")
            return

        # 命中可接班条件（全部可接 或 部分可接）！立即触发秒回！
        final_reply = res.suggested_reply or reply_word
        self._append_log(f"🎯 发现符合条件的替班！{sdr_disp}{preview}\n   -> {res.decision_reason}", "TAKE")
        sent_ok, sent_msg = self.monitor.send_reply_to_wechat(final_reply)
        self._append_log(f"   -> 🚀 {sent_msg}", "TAKE" if sent_ok else "ERR")

        if sent_ok and self.auto_update_var.get():
            for item in res.accepted_items:
                ok, sub_msg, _ = self.schedule_mgr.apply_substitution(
                    slots_to_take=item.matched_slots,
                    operator_name=self.schedule_mgr.operator_name,
                    requester_name=item.detected_requester,
                    source_message=f"{sdr_disp}{preview}",
                    auto_replied=True,
                )
                if ok:
                    self._append_log(f"   -> 📅 已更新个人班表：{sub_msg}", "TAKE")
            self._refresh_all_views()

    def _append_log(self, text: str, tag: str = "INFO") -> None:
        self.log_text.insert(tk.END, text + "\n", tag)
        self.log_text.see(tk.END)

    def _on_close(self) -> None:
        if self.monitor.is_running:
            self.monitor.stop()
        self.schedule_mgr.save_state()
        self.parser.save_rules()
        self.destroy()


if __name__ == "__main__":
    app = LibraryShiftApp()
    app.mainloop()
