# -*- coding: utf-8 -*-
"""
图书馆助理排班与替班管理核心模块 (schedule_manager.py)
负责：
1. 读取解析每周固定排班表 (.xlsx)，支持服务台、2/3/4楼书库、D5图书室全部5个工作表
2. 提取所有助理名单、个人当前排班、多人同班次槽位 (如 18:00-20:00 双人岗)
3. 同步实时日期/星期几，并将排班表标题中的日期范围 (如 11月16日-11月22日) 映射到周一至周日
4. 管理个人“已有班次”、“不可接时段/班次”、“可接班次”筛选与时间冲突检测 (如 20:00-22:00 与 20:00-23:00 重叠互斥)
5. 支持连班组合匹配 (如 14:00-18:00 拆分为 14:00-15:30 + 15:30-18:00，18:00-22:00 拆分为 18:00-20:00 + 20:00-22:00)
6. 替班成功后精准更新个人班次（多人单元格仅替换被替班人姓名，绝不搞混别人的班），支持一键撤销与导出最新 Excel 排班表
"""

from __future__ import annotations
import copy
import datetime
import json
import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional, Set, Tuple

WEEKDAYS = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
WEEKDAY_MAP = {
    "周一": 0, "星期一": 0, "礼拜一": 0, "礼拜1": 0, "周1": 0,
    "周二": 1, "星期二": 1, "礼拜二": 1, "礼拜2": 1, "周2": 1,
    "周三": 2, "星期三": 2, "礼拜三": 2, "礼拜3": 2, "周3": 2,
    "周四": 3, "星期四": 3, "礼拜四": 3, "礼拜4": 3, "周4": 3,
    "周五": 4, "星期五": 4, "礼拜五": 4, "礼拜5": 4, "周5": 4,
    "周六": 5, "星期六": 5, "礼拜六": 5, "礼拜6": 5, "周6": 5,
    "周日": 6, "周天": 6, "星期日": 6, "星期天": 6, "礼拜日": 6, "礼拜天": 6,
}

# 标准时段定义（用于全周可视化排班矩阵展示与时间段映射）
STANDARD_TIME_SLOTS = [
    ("08:00-09:00", 480, 540, "平日早班 (服务台)"),
    ("08:00-09:30", 480, 570, "周末早班 (服务台)"),
    ("09:00-10:30", 540, 630, "平日上午一节"),
    ("09:00-12:00", 540, 720, "周末上午 (书库)"),
    ("09:30-12:00", 570, 720, "周末上午 (服务台)"),
    ("10:30-12:30", 630, 750, "平日上午二节"),
    ("12:00-14:00", 720, 840, "周末午班 (服务台)"),
    ("12:30-14:00", 750, 840, "平日午班 (服务台)"),
    ("14:00-15:30", 840, 930, "平日下午一节"),
    ("14:00-16:00", 840, 960, "周末下午一节"),
    ("15:30-18:00", 930, 1080, "平日下午二节"),
    ("16:00-18:00", 960, 1080, "周末下午二节"),
    ("18:00-20:00", 1080, 1200, "晚班一节 (18-20点)"),
    ("20:00-22:00", 1200, 1320, "晚班A岗 (20-22点/仅书库)"),
    ("20:00-23:00", 1200, 1380, "晚班B/闭馆岗 (20-23点)"),
    ("上午", -1, -1, "D5图书室上午 (弹性1h)"),
    ("晚上", -2, -2, "D5图书室晚上 (弹性1h)"),
]


def normalize_post_name(sheet_name: str) -> str:
    """规范化工作表/岗位名称"""
    s = sheet_name.strip()
    if "服务台" in s:
        return "服务台"
    if "2楼" in s or "二楼" in s:
        return "2楼书库"
    if "3楼" in s or "三楼" in s:
        return "3楼书库"
    if "4楼" in s or "四楼" in s:
        return "4楼书库"
    if "D5" in s.upper():
        return "D5图书室"
    return s


def normalize_time_str(raw_time: str) -> Tuple[str, int, int]:
    """
    将排班表中的时段字符串标准化，如 '8:00-9：00' -> ('08:00-09:00', 480, 540)
    对于 D5 图书室的 '上午' / '晚上'，返回 ('上午', -1, -1) / ('晚上', -2, -2)
    """
    t = str(raw_time).strip().replace("：", ":").replace(" ", "").replace("—", "-").replace("~", "-").replace("至", "-")
    if "上午" in t:
        return "上午", -1, -1
    if "晚上" in t:
        return "晚上", -2, -2

    m = re.match(r"^(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})$", t)
    if m:
        h1, m1, h2, m2 = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
        norm = f"{h1:02d}:{m1:02d}-{h2:02d}:{m2:02d}"
        return norm, h1 * 60 + m1, h2 * 60 + m2
    return t, 0, 0


def split_names(raw_names: str) -> List[str]:
    """拆分单元格中的姓名（支持多人同班次，如 '张三，李四' 或 '王五、赵六'）"""
    if not raw_names:
        return []
    parts = re.split(r"[,，、/；;\s]+", str(raw_names).strip())
    return [p.strip() for p in parts if p.strip()]


def join_names(names: List[str], original_raw: str = "") -> str:
    """将多人姓名重新拼接回单元格字符串，尽量保留原分隔符风格"""
    if not names:
        return ""
    if len(names) == 1:
        return names[0]
    sep = "，"
    if "、" in original_raw:
        sep = "、"
    elif "," in original_raw:
        sep = ","
    return sep.join(names)


def times_overlap(start1: int, end1: int, start2: int, end2: int) -> bool:
    """
    判断两个时间段（以分钟计）是否重叠。
    注意：D5图书室的 '上午'(-1) 和 '晚上'(-2) 是弹性打卡1小时，默认仅与同为D5同半天的班次冲突，不与固定坐班硬性互斥（可通过配置调整）。
    """
    if start1 < 0 or start2 < 0:
        return start1 == start2 and start1 < 0
    return max(start1, start2) < min(end1, end2)


@dataclass
class ShiftSlot:
    """代表排班表中的一个具体班次单元格"""
    slot_id: str                 # 唯一标识，如 "服务台|周一|08:00-09:00"
    post: str                    # 规范化岗位名：服务台 / 2楼书库 / 3楼书库 / 4楼书库 / D5图书室
    sheet_name: str              # Excel 原工作表名（可能带尾部空格）
    cell_ref: str                # 姓名所在单元格坐标，如 "C3", "H7"
    weekday: str                 # 周一 ~ 周日
    weekday_idx: int             # 0 ~ 6
    date_str: str                # 对应日期，如 "11月16日" (若可解析)
    iso_date: str                # 对应 ISO 日期，如 "2026-11-16"
    time_str: str                # 标准化时间段，如 "08:00-09:00", "20:00-23:00", "上午"
    raw_time_str: str            # Excel 原始时间段文本
    start_min: int               # 开始分钟数 (如 08:00 -> 480)
    end_min: int                 # 结束分钟数 (如 09:00 -> 540)
    original_raw_names: str      # Excel 原始姓名字符串
    original_holders: List[str]  # 初始排班负责人列表
    current_holders: List[str]   # 当前排班负责人列表（替班后更新）
    extra_info: str = ""         # 附加信息（如 D5 的书号范围与工作内容）

    def has_holder(self, name: str) -> bool:
        return name in self.current_holders

    def was_originally_held_by(self, name: str) -> bool:
        return name in self.original_holders

    def display_holders(self) -> str:
        return "、".join(self.current_holders)


@dataclass
class SubstitutionRecord:
    """替班历史记录（用于追踪替班明细与一键撤销）"""
    record_id: str
    timestamp: str               # 记录时间
    operator_name: str           # 接班人（操作者）
    original_person: str         # 被替班人（原班主）
    post: str                    # 岗位
    weekday: str                 # 星期几
    date_str: str                # 日期
    time_str: str                # 时间段（可能是单班或连班描述）
    slot_ids: List[str]          # 涉及的 slot_id 列表
    previous_holders_map: Dict[str, List[str]]  # 替班前各 slot_id 的 current_holders 备份
    source_message: str = ""     # 触发该替班的微信原消息
    auto_replied: bool = False   # 是否由监控自动回复触发
    status: str = "已生效"        # "已生效" / "已撤销"


class ScheduleManager:
    """排班表解析、个人筛选配置与替班状态管理器"""

    def __init__(self, excel_path: str, state_path: str = "schedule_state.json"):
        self.excel_path = os.path.abspath(excel_path)
        self.state_path = os.path.abspath(state_path)
        self.week_title: str = ""
        self.weekday_dates: Dict[str, Tuple[str, str]] = {}  # "周一" -> ("11月16日", "2026-11-16")
        self.slots: Dict[str, ShiftSlot] = {}
        self.slots_order: List[str] = []
        self.all_names: List[str] = []

        # 用户配置状态
        self.operator_name: str = ""
        # 用户手动标记“不能接”的时间格子：存 "周一|08:00-09:00" 形式的 key（对所有岗位生效的时间冲突）
        self.blocked_time_keys: Set[str] = set()
        # 用户手动标记“不能接”的特定岗位班次 slot_id
        self.blocked_slot_ids: Set[str] = set()
        # 用户启用/允许接班的岗位集合
        self.allowed_posts: Set[str] = {"服务台", "2楼书库", "3楼书库", "4楼书库", "D5图书室"}
        # 替班记录列表
        self.records: List[SubstitutionRecord] = []
        # 监控与回复设置
        self.reply_text: str = "我可以"
        self.target_group_name: str = ""
        self.ignore_expired_shifts: bool = False
        self.auto_update_on_reply: bool = True
        self.auto_sync_week_dates: bool = True
        self._excel_date_source: str = ""

        if os.path.exists(self.excel_path):
            self.load_excel(self.excel_path)
        self.load_state()
        if self.auto_sync_week_dates:
            self.sync_dates_to_current_week(use_realtime=True)

    # ------------------------------------------------------------------
    # 1. 解析 Excel 排班表
    # ------------------------------------------------------------------
    def _parse_week_dates(self, title_or_filename: str, ref_year: Optional[int] = None, allow_fallback: bool = True) -> bool:
        """从文件名或表头（如 '11月16日-11月22日'）解析周一到周日的具体日期，返回是否显式匹配到日期范围"""
        if ref_year is None:
            ref_year = datetime.datetime.now().year
        m = re.search(r"(\d{1,2})月(\d{1,2})日\s*[-~至—]\s*(\d{1,2})月(\d{1,2})日", title_or_filename)
        if m:
            m1, d1, m2, d2 = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
            self.week_title = f"{m1}月{d1}日-{m2}月{d2}日"
            try:
                start_dt = datetime.date(ref_year, m1, d1)
                for i, wd in enumerate(WEEKDAYS):
                    cur = start_dt + datetime.timedelta(days=i)
                    self.weekday_dates[wd] = (f"{cur.month}月{cur.day}日", cur.isoformat())
                return True
            except ValueError:
                pass

        if not allow_fallback:
            return False

        # 若未匹配到，则按当前日期所在周计算
        today = datetime.date.today()
        monday = today - datetime.timedelta(days=today.weekday())
        self.week_title = f"{monday.month}月{monday.day}日-{(monday + datetime.timedelta(days=6)).month}月{(monday + datetime.timedelta(days=6)).day}日"
        for i, wd in enumerate(WEEKDAYS):
            cur = monday + datetime.timedelta(days=i)
            self.weekday_dates[wd] = (f"{cur.month}月{cur.day}日", cur.isoformat())
        return False

    def sync_dates_to_current_week(self, use_realtime: Optional[bool] = None) -> None:
        """
        根据实时日历或原表格文件名同步本周周一至周日的日期，并刷新所有班次对应的日期字段。
        这样即便使用同一份固定排班模板跨周循环，周一到周日的日期也会自动替换成当前周真实日期！
        """
        if use_realtime is not None:
            self.auto_sync_week_dates = bool(use_realtime)

        if self.auto_sync_week_dates or not self._excel_date_source:
            today = datetime.date.today()
            monday = today - datetime.timedelta(days=today.weekday())
            self.week_title = f"{monday.month}月{monday.day}日-{(monday + datetime.timedelta(days=6)).month}月{(monday + datetime.timedelta(days=6)).day}日"
            for i, wd in enumerate(WEEKDAYS):
                cur = monday + datetime.timedelta(days=i)
                self.weekday_dates[wd] = (f"{cur.month}月{cur.day}日", cur.isoformat())
        else:
            self._parse_week_dates(self._excel_date_source, allow_fallback=True)

        for slot in self.slots.values():
            d_str, iso_d = self.weekday_dates.get(slot.weekday, ("", ""))
            slot.date_str = d_str
            slot.iso_date = iso_d

    def load_excel(self, excel_path: str) -> None:
        """使用内置 zipfile + xml 解析 .xlsx，零外部依赖也可完整读取所有工作表与坐标"""
        self.excel_path = os.path.abspath(excel_path)
        self.slots.clear()
        self.slots_order.clear()
        self.week_title = ""
        self._excel_date_source = ""
        matched_date_range = self._parse_week_dates(os.path.basename(self.excel_path), allow_fallback=True)
        if matched_date_range:
            self._excel_date_source = os.path.basename(self.excel_path)

        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
              "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

        with zipfile.ZipFile(self.excel_path, "r") as z:
            shared: List[str] = []
            if "xl/sharedStrings.xml" in z.namelist():
                root = ET.fromstring(z.read("xl/sharedStrings.xml"))
                for si in root:
                    texts = [t.text or "" for t in si.iter(f"{{{ns['main']}}}t")]
                    shared.append("".join(texts))

            wb = ET.fromstring(z.read("xl/workbook.xml"))
            rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            rid_to_target = {}
            for rel in rels:
                rid_to_target[rel.attrib["Id"]] = rel.attrib["Target"]

            sheet_entries = []
            for s in wb.iter(f"{{{ns['main']}}}sheet"):
                s_name = s.attrib["name"]
                rid = s.attrib[f"{{{ns['r']}}}id"]
                target = rid_to_target.get(rid, "").lstrip("/")
                if not target.startswith("xl/"):
                    target = "xl/" + target
                sheet_entries.append((s_name, target))

            for raw_sheet_name, target_path in sheet_entries:
                if target_path not in z.namelist():
                    continue
                ws_xml = ET.fromstring(z.read(target_path))
                grid: Dict[Tuple[int, str], str] = {}
                max_row = 0
                for row_el in ws_xml.iter(f"{{{ns['main']}}}row"):
                    r_idx = int(row_el.attrib.get("r", "0"))
                    max_row = max(max_row, r_idx)
                    for c_el in row_el:
                        ref = c_el.attrib.get("r", "")
                        col_m = re.match(r"^([A-Z]+)(\d+)$", ref)
                        if not col_m:
                            continue
                        col_str = col_m.group(1)
                        t_type = c_el.attrib.get("t", "")
                        v_el = c_el.find(f"{{{ns['main']}}}v")
                        val = v_el.text if v_el is not None and v_el.text is not None else ""
                        if t_type == "s" and val != "":
                            val = shared[int(val)]
                        elif t_type == "inlineStr":
                            t_els = [t.text or "" for t in c_el.iter(f"{{{ns['main']}}}t")]
                            val = "".join(t_els)
                        if val != "":
                            grid[(r_idx, col_str)] = val.strip()

                if (1, "A") in grid and not matched_date_range:
                    if self._parse_week_dates(grid[(1, "A")], allow_fallback=False):
                        matched_date_range = True
                        self._excel_date_source = grid[(1, "A")]

                post = normalize_post_name(raw_sheet_name)
                if post == "D5图书室":
                    self._parse_d5_sheet(post, raw_sheet_name, grid, max_row)
                elif post == "服务台":
                    # 服务台：左栏 A(日期), B(时段), C(姓名)；右栏 E(日期), F(时段), G(姓名)
                    self._parse_two_column_sheet(
                        post, raw_sheet_name, grid, max_row,
                        left_cols=("A", "B", "C"),
                        right_cols=("E", "F", "G")
                    )
                else:
                    # 2/3/4楼书库：左栏 A(日期), B(时段), C(姓名)；右栏 F(日期), G(时段), H(姓名)
                    self._parse_two_column_sheet(
                        post, raw_sheet_name, grid, max_row,
                        left_cols=("A", "B", "C"),
                        right_cols=("F", "G", "H")
                    )

        # 汇总所有姓名并按拼音/出现顺序排序
        names_set: Set[str] = set()
        for slot in self.slots.values():
            for n in slot.original_holders:
                names_set.add(n)
            for n in slot.current_holders:
                names_set.add(n)
        self.all_names = sorted(names_set)

    def _parse_two_column_sheet(
        self,
        post: str,
        raw_sheet_name: str,
        grid: Dict[Tuple[int, str], str],
        max_row: int,
        left_cols: Tuple[str, str, str],
        right_cols: Tuple[str, str, str],
    ) -> None:
        """解析双栏排班表（服务台、2楼书库、3楼书库、4楼书库）"""
        for day_col, time_col, name_col in [left_cols, right_cols]:
            current_weekday = ""
            for r in range(1, max_row + 1):
                day_val = grid.get((r, day_col), "").strip()
                if day_val in WEEKDAY_MAP:
                    current_weekday = WEEKDAYS[WEEKDAY_MAP[day_val]]
                elif day_val in ("日期", "星期") or "签到" in day_val:
                    current_weekday = ""
                    continue

                time_val = grid.get((r, time_col), "").strip()
                name_val = grid.get((r, name_col), "").strip()
                if not current_weekday or not time_val or time_val == "时段":
                    continue
                if not name_val or name_val == "姓名":
                    continue

                norm_time, start_min, end_min = normalize_time_str(time_val)
                if start_min == 0 and end_min == 0:
                    continue

                wd_idx = WEEKDAY_MAP[current_weekday]
                date_str, iso_date = self.weekday_dates.get(current_weekday, ("", ""))
                slot_id = f"{post}|{current_weekday}|{norm_time}"
                holders = split_names(name_val)
                slot = ShiftSlot(
                    slot_id=slot_id,
                    post=post,
                    sheet_name=raw_sheet_name,
                    cell_ref=f"{name_col}{r}",
                    weekday=current_weekday,
                    weekday_idx=wd_idx,
                    date_str=date_str,
                    iso_date=iso_date,
                    time_str=norm_time,
                    raw_time_str=time_val,
                    start_min=start_min,
                    end_min=end_min,
                    original_raw_names=name_val,
                    original_holders=list(holders),
                    current_holders=list(holders),
                )
                self.slots[slot_id] = slot
                self.slots_order.append(slot_id)

        # 对该工作表的班次按 (weekday_idx, start_min, end_min) 排序，确保顺序整洁
        self.slots_order.sort(
            key=lambda sid: (
                ["服务台", "2楼书库", "3楼书库", "4楼书库", "D5图书室"].index(self.slots[sid].post)
                if self.slots[sid].post in ["服务台", "2楼书库", "3楼书库", "4楼书库", "D5图书室"] else 99,
                self.slots[sid].weekday_idx,
                self.slots[sid].start_min if self.slots[sid].start_min >= 0 else (600 if self.slots[sid].start_min == -1 else 1140),
                self.slots[sid].end_min,
            )
        )

    def _parse_d5_sheet(
        self,
        post: str,
        raw_sheet_name: str,
        grid: Dict[Tuple[int, str], str],
        max_row: int,
    ) -> None:
        """解析 D5图书室值班表"""
        current_weekday = ""
        for r in range(3, max_row + 1):
            day_val = grid.get((r, "A"), "").strip()
            if day_val in WEEKDAY_MAP:
                current_weekday = WEEKDAYS[WEEKDAY_MAP[day_val]]
            elif "说明" in day_val:
                break

            time_val = grid.get((r, "B"), "").strip()
            name_val = grid.get((r, "C"), "").strip()
            area_val = grid.get((r, "E"), "").strip().replace("\n", " ")
            if not current_weekday or not time_val or not name_val:
                continue

            norm_time, start_min, end_min = normalize_time_str(time_val)
            wd_idx = WEEKDAY_MAP[current_weekday]
            date_str, iso_date = self.weekday_dates.get(current_weekday, ("", ""))
            slot_id = f"{post}|{current_weekday}|{norm_time}"
            holders = split_names(name_val)
            slot = ShiftSlot(
                slot_id=slot_id,
                post=post,
                sheet_name=raw_sheet_name,
                cell_ref=f"C{r}",
                weekday=current_weekday,
                weekday_idx=wd_idx,
                date_str=date_str,
                iso_date=iso_date,
                time_str=norm_time,
                raw_time_str=time_val,
                start_min=start_min,
                end_min=end_min,
                original_raw_names=name_val,
                original_holders=list(holders),
                current_holders=list(holders),
                extra_info=area_val,
            )
            self.slots[slot_id] = slot
            self.slots_order.append(slot_id)

    # ------------------------------------------------------------------
    # 2. 状态持久化（保存/读取用户筛选与替班记录）
    # ------------------------------------------------------------------
    def save_state(self) -> None:
        """保存当前操作者配置、禁接班次和替班记录到 JSON"""
        current_holders_override = {}
        for sid, slot in self.slots.items():
            if slot.current_holders != slot.original_holders:
                current_holders_override[sid] = slot.current_holders

        data = {
            "excel_path": self.excel_path,
            "operator_name": self.operator_name,
            "blocked_time_keys": sorted(self.blocked_time_keys),
            "blocked_slot_ids": sorted(self.blocked_slot_ids),
            "allowed_posts": sorted(self.allowed_posts),
            "reply_text": self.reply_text,
            "target_group_name": self.target_group_name,
            "ignore_expired_shifts": self.ignore_expired_shifts,
            "auto_update_on_reply": self.auto_update_on_reply,
            "auto_sync_week_dates": self.auto_sync_week_dates,
            "current_holders_override": current_holders_override,
            "records": [asdict(r) for r in self.records],
        }
        with open(self.state_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def load_state(self) -> None:
        """从 JSON 恢复保存的筛选配置和替班状态"""
        if not os.path.exists(self.state_path):
            return
        try:
            with open(self.state_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            saved_excel = data.get("excel_path", "")
            if not self.slots and saved_excel and os.path.exists(saved_excel):
                self.load_excel(saved_excel)

            self.operator_name = data.get("operator_name", "")
            self.blocked_time_keys = set(data.get("blocked_time_keys", []))
            self.blocked_slot_ids = set(data.get("blocked_slot_ids", []))
            if "allowed_posts" in data:
                self.allowed_posts = set(data["allowed_posts"])
            self.reply_text = data.get("reply_text", "我可以")
            self.target_group_name = data.get("target_group_name", "")
            # 按周期排班逻辑不再自动屏蔽“已过时间”（有人会提前发下周同天班次求替）
            self.ignore_expired_shifts = False
            self.auto_update_on_reply = data.get("auto_update_on_reply", True)
            self.auto_sync_week_dates = data.get("auto_sync_week_dates", True)

            same_excel = (
                not saved_excel
                or os.path.abspath(saved_excel) == os.path.abspath(self.excel_path)
                or os.path.basename(saved_excel) == os.path.basename(self.excel_path)
            )
            if same_excel:
                overrides = data.get("current_holders_override", {})
                for sid, holders in overrides.items():
                    if sid in self.slots:
                        self.slots[sid].current_holders = list(holders)
                self.records = [SubstitutionRecord(**r) for r in data.get("records", [])]
        except Exception as e:
            print(f"[警告] 读取状态文件失败: {e}")

    def reset_week_substitutions(self) -> None:
        """重置本周的所有替班变动，恢复为 Excel 初始固定排班"""
        for slot in self.slots.values():
            slot.current_holders = list(slot.original_holders)
        self.records.clear()
        self.save_state()

    # ------------------------------------------------------------------
    # 3. 个人排班查询、冲突检测与可接班次计算
    # ------------------------------------------------------------------
    def get_person_current_slots(self, name: str) -> List[ShiftSlot]:
        """获取某人当前拥有的所有班次（包含原固定班次与替班接下的班次，不含已转出的班次）"""
        if not name:
            return []
        return [
            self.slots[sid]
            for sid in self.slots_order
            if self.slots[sid].has_holder(name)
        ]

    def get_person_original_slots(self, name: str) -> List[ShiftSlot]:
        """获取某人最初固定排班的所有班次"""
        if not name:
            return []
        return [
            self.slots[sid]
            for sid in self.slots_order
            if self.slots[sid].was_originally_held_by(name)
        ]

    def is_time_window_conflicting_with_existing(self, name: str, weekday: str, start_min: int, end_min: int) -> Optional[ShiftSlot]:
        """
        检查某人在指定星期与时间段是否已有班次冲突（例如已有 20:00-22:00 的班，则不能再接 20:00-23:00 的班）。
        若冲突，返回冲突的已有 ShiftSlot；否则返回 None。
        """
        if not name:
            return None
        for slot in self.get_person_current_slots(name):
            if slot.weekday == weekday and times_overlap(slot.start_min, slot.end_min, start_min, end_min):
                return slot
        return None

    def get_post_time_rows(self, post: str) -> List[Tuple[str, int, int, str]]:
        """获取指定岗位在排班表中实际存在的所有去重时段行（例如服务台无20:00-22:00，书库含20-22与20-23两班岗）"""
        desc_map = {t: desc for t, _, _, desc in STANDARD_TIME_SLOTS}
        found: Dict[str, Tuple[int, int]] = {}
        for s in self.slots.values():
            if s.post == post and s.time_str not in found:
                found[s.time_str] = (s.start_min, s.end_min)

        rows: List[Tuple[str, int, int, str]] = []
        for t_str, (st, ed) in found.items():
            desc = desc_map.get(t_str, "")
            if not desc:
                if post == "D5图书室":
                    desc = "D5弹性1小时"
                elif st >= 0:
                    desc = post
            elif post == "服务台":
                desc = desc.replace(" (服务台)", "").replace(" (20-23点)", "")
                if t_str == "20:00-23:00":
                    desc = "晚班闭馆 (服务台无20-22)"
            elif "书库" in post:
                desc = desc.replace(" (书库)", "").replace("/仅书库", "")
                if t_str == "18:00-20:00":
                    desc = "晚班一节 (双人岗)"
                elif t_str == "20:00-22:00":
                    desc = "晚班两班岗之一 (20-22点)"
                elif t_str == "20:00-23:00":
                    desc = "晚班两班岗之一 (20-23点)"
            rows.append((t_str, st, ed, desc))

        rows.sort(key=lambda r: (0 if r[1] >= 0 else 1, r[1] if r[1] >= 0 else -r[1], r[2], r[0]))
        return rows

    def is_time_window_blocked_by_user(self, weekday: str, start_min: int, end_min: int, time_str: str = "") -> bool:
        """
        兼容旧版全局禁接键（精确匹配 weekday|time_str，不再做区间重叠误杀，保证 20-22 与 20-23 互不牵连）。
        新界面按岗位独立将 slot_id 存入 blocked_slot_ids。
        """
        if time_str and f"{weekday}|{time_str}" in self.blocked_time_keys:
            return True
        return False

    def is_slot_expired(self, slot: ShiftSlot, now: Optional[datetime.datetime] = None) -> bool:
        """判断班次相对于当前实时时间是否已经过期"""
        if now is None:
            now = datetime.datetime.now()
        today_iso = now.date().isoformat()
        if slot.iso_date:
            if slot.iso_date < today_iso:
                return True
            if slot.iso_date > today_iso:
                return False
            # 如果是今天，检查班次结束时间（或开始后30分钟）是否已过
            if slot.start_min > 0:
                now_min = now.hour * 60 + now.minute
                # 班次已经开始超过15分钟或已结束则视为过期
                if now_min > slot.start_min + 15:
                    return True
        return False

    def evaluate_slot_status(
        self,
        slot: ShiftSlot,
        operator_name: Optional[str] = None,
        check_expired: bool = False,
        now: Optional[datetime.datetime] = None,
    ) -> Tuple[str, str]:
        """
        评估某个具体班次对当前操作者的状态：
        返回值：(status_code, reason)
        status_code 取值：
          - "OWNED": 操作者自己目前已有的班次
          - "CONFLICT": 与操作者同一天同时间段/重叠时间段的已有班次冲突
          - "BLOCKED_POST": 该岗位已被用户全局取消勾选（不接该岗位）
          - "BLOCKED_TIME": 该时间段（或重叠时间段）已被用户标记为不可接
          - "BLOCKED_SLOT": 该具体岗位班次被用户单独标记为不可接
          - "EXPIRED": 该班次时间已过
          - "AVAILABLE": 可以接班！
        """
        op = operator_name if operator_name is not None else self.operator_name
        if op and slot.has_holder(op):
            is_orig = slot.was_originally_held_by(op)
            return "OWNED", ("本人原固定班次" if is_orig else "本人已替班接下")

        if op:
            conflict_slot = self.is_time_window_conflicting_with_existing(op, slot.weekday, slot.start_min, slot.end_min)
            if conflict_slot is not None:
                return "CONFLICT", f"与本人已有班次冲突 ({conflict_slot.post} {conflict_slot.time_str})"

        if slot.post not in self.allowed_posts:
            return "BLOCKED_POST", f"已设置不接【{slot.post}】岗位"

        if slot.slot_id in self.blocked_slot_ids:
            return "BLOCKED_SLOT", "已单独设为不可接班次"

        if self.is_time_window_blocked_by_user(slot.weekday, slot.start_min, slot.end_min, slot.time_str):
            return "BLOCKED_TIME", "处于个人设为不可接的时间段"

        if check_expired and self.ignore_expired_shifts and self.is_slot_expired(slot, now=now):
            return "EXPIRED", "该班次时间已过"

        return "AVAILABLE", "符合条件，可接班"

    # ------------------------------------------------------------------
    # 4. 根据替班请求智能匹配班次（支持单班、连班、按姓名反查岗位）
    # ------------------------------------------------------------------
    def find_matching_slots_for_request(
        self,
        weekday: str,
        start_min: int,
        end_min: int,
        raw_time_token: str = "",
        post: Optional[str] = None,
        requester_name: Optional[str] = None,
    ) -> List[List[ShiftSlot]]:
        """
        根据解析出的替班要素（星期、起止时间、可选岗位、可选发信人姓名），在排班表中查找匹配的班次组合。
        支持：
        1. 精确单班匹配（如 周一 14:00-15:30 服务台）
        2. 连续多班组合匹配（如 周一 14:00-18:00 服务台 -> 匹配 [14:00-15:30, 15:30-18:00]；
           周三 18:00-22:00 3楼书库 -> 匹配 [18:00-20:00, 20:00-22:00]；
           周三 18:00-23:00 3楼书库 -> 匹配 [18:00-20:00, 20:00-23:00]）
        3. 若未说明具体岗位，但提供了发信人姓名（或群昵称包含真名），自动反查此人在该时段所在的岗位！
        返回候选的班次组合列表 List[List[ShiftSlot]]。
        """
        candidate_posts = [post] if post else ["服务台", "2楼书库", "3楼书库", "4楼书库", "D5图书室"]
        results: List[List[ShiftSlot]] = []

        for p in candidate_posts:
            day_slots = [
                self.slots[sid]
                for sid in self.slots_order
                if self.slots[sid].post == p and self.slots[sid].weekday == weekday
            ]
            if not day_slots:
                continue

            # 情况A：D5图书室（上午/晚上）
            if p == "D5图书室":
                for s in day_slots:
                    if raw_time_token == s.time_str or (start_min < 0 and s.start_min == start_min):
                        results.append([s])
                continue

            # 情况B：精确匹配单个班次
            exact = [s for s in day_slots if s.start_min == start_min and s.end_min == end_min]
            if exact:
                results.append(exact)
                continue

            # 情况C：连班组合匹配（例如 14:00-18:00 = 14:00-15:30 + 15:30-18:00，或 18:00-23:00 = 18:00-20:00 + 20:00-23:00）
            if start_min > 0 and end_min > start_min:
                chain = self._find_contiguous_chain(day_slots, start_min, end_min, requester_name)
                if chain:
                    results.append(chain)

        # 如果提供了发信人姓名（如微信群昵称里带有排班表中的姓名），优先筛选原班主或当前班主包含该姓名的结果！
        if requester_name and results:
            matched_by_name = [
                combo for combo in results
                if all(s.has_holder(requester_name) or s.was_originally_held_by(requester_name) for s in combo)
            ]
            if matched_by_name:
                return matched_by_name

        return results

    def _find_contiguous_chain(
        self,
        day_slots: List[ShiftSlot],
        target_start: int,
        target_end: int,
        requester_name: Optional[str] = None,
    ) -> Optional[List[ShiftSlot]]:
        """使用 DFS 在同一岗位同一天的班次中寻找首尾相接覆盖 [target_start, target_end] 的连班链"""
        valid_chains: List[List[ShiftSlot]] = []

        def dfs(cur_min: int, path: List[ShiftSlot]):
            if cur_min == target_end:
                valid_chains.append(list(path))
                return
            for s in day_slots:
                if s.start_min == cur_min and s.end_min <= target_end and s not in path:
                    path.append(s)
                    dfs(s.end_min, path)
                    path.pop()

        dfs(target_start, [])
        if not valid_chains:
            return None
        if requester_name:
            for ch in valid_chains:
                if all(s.has_holder(requester_name) for s in ch):
                    return ch
        return valid_chains[0]

    # ------------------------------------------------------------------
    # 5. 执行替班更新、撤销替班与导出 Excel
    # ------------------------------------------------------------------
    def apply_substitution(
        self,
        slots_to_take: List[ShiftSlot],
        operator_name: Optional[str] = None,
        requester_name: str = "",
        source_message: str = "",
        auto_replied: bool = False,
    ) -> Tuple[bool, str, Optional[SubstitutionRecord]]:
        """
        将指定班次（单班或连班）更新到操作者名下，同时确保绝不覆盖同单元格的其他人！
        规则：
        - 若单元格只有1人（如 '张三'），替换为操作者姓名。
        - 若单元格有2人（如 18:00-20:00 的 '李四、王五'）：
          1) 若已知 requester_name 在其中，精准替换 requester_name；
          2) 若属于连班（如 18:00-22:00），自动对比下一节（20:00-22:00 是 '李四'），从而精准推断出是替 '李四' 的班，保留 '王五'！
          3) 若仍无法区分且两人都在，默认替换第一位并在记录中标注被替班人。
        """
        op = (operator_name or self.operator_name).strip()
        if not op:
            return False, "未设置操作者姓名，无法更新排班", None
        if not slots_to_take:
            return False, "未指定要接的班次", None

        # 自动推断连班中的共同原班主（针对 18:00-20:00 双人岗）
        inferred_requester = requester_name.strip()
        if not inferred_requester and len(slots_to_take) > 1:
            common = set(slots_to_take[0].current_holders)
            for s in slots_to_take[1:]:
                common &= set(s.current_holders)
            common.discard(op)
            if len(common) == 1:
                inferred_requester = next(iter(common))

        prev_map: Dict[str, List[str]] = {}
        replaced_people: List[str] = []

        for s in slots_to_take:
            if op in s.current_holders:
                return False, f"您已在【{s.post} {s.weekday} {s.time_str}】名单中，无需重复替班", None

            prev_map[s.slot_id] = list(s.current_holders)
            new_holders = list(s.current_holders)

            if len(new_holders) <= 1:
                if new_holders:
                    replaced_people.append(new_holders[0])
                new_holders = [op]
            else:
                # 多人同班次单元格，精准替换目标人物，绝不搞混同伴的班
                target_idx = 0
                if inferred_requester and inferred_requester in new_holders:
                    target_idx = new_holders.index(inferred_requester)
                replaced_people.append(new_holders[target_idx])
                new_holders[target_idx] = op

            s.current_holders = new_holders

        unique_replaced = "、".join(dict.fromkeys(replaced_people)) if replaced_people else (inferred_requester or "原班主")
        first = slots_to_take[0]
        last = slots_to_take[-1]
        if len(slots_to_take) == 1:
            time_desc = first.time_str
        else:
            t_start = first.time_str.split("-")[0]
            t_end = last.time_str.split("-")[-1]
            time_desc = f"{t_start}-{t_end} (含{len(slots_to_take)}个连班)"

        rec = SubstitutionRecord(
            record_id=f"SUB-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}-{len(self.records)+1}",
            timestamp=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            operator_name=op,
            original_person=unique_replaced,
            post=first.post,
            weekday=first.weekday,
            date_str=first.date_str,
            time_str=time_desc,
            slot_ids=[s.slot_id for s in slots_to_take],
            previous_holders_map=prev_map,
            source_message=source_message,
            auto_replied=auto_replied,
            status="已生效",
        )
        self.records.append(rec)
        if op not in self.all_names:
            self.all_names.append(op)
            self.all_names.sort()
        self.save_state()
        return True, f"已成功接替【{unique_replaced}】的 {first.weekday}({first.date_str}) {first.post} {time_desc} 班次！", rec

    def undo_substitution(self, record_id: str) -> Tuple[bool, str]:
        """一键撤销某次替班记录（例如微信群里慢了一步被别人抢先时，一键恢复原班表）"""
        for rec in reversed(self.records):
            if rec.record_id == record_id:
                if rec.status == "已撤销":
                    return False, "该替班记录已经处于已撤销状态"
                for sid, prev_holders in rec.previous_holders_map.items():
                    if sid in self.slots:
                        self.slots[sid].current_holders = list(prev_holders)
                rec.status = "已撤销"
                self.save_state()
                return True, f"已撤销【{rec.weekday} {rec.post} {rec.time_str}】的替班，排班表已恢复原状。"
        return False, "未找到对应的替班记录"

    def export_updated_excel(self, output_path: Optional[str] = None) -> Tuple[bool, str]:
        """
        将更新后的排班导出到 Excel 文件（保留原表格全部格式、合并单元格与边框）。
        默认保存到原目录下的 '<原文件名>_个人最新排班.xlsx'，并提供原文件备份。
        """
        if output_path is None:
            base, ext = os.path.splitext(self.excel_path)
            output_path = f"{base}_最新排班{ext}"

        try:
            import openpyxl
            shutil.copy2(self.excel_path, output_path)
            wb = openpyxl.load_workbook(output_path)
            if self.week_title:
                for ws in wb.worksheets:
                    for row_idx in (1, 24, 25):
                        val = ws.cell(row_idx, 1).value
                        if isinstance(val, str) and re.search(r"\d{1,2}月\d{1,2}日\s*[-~至—]\s*\d{1,2}月\d{1,2}日", val):
                            ws.cell(row_idx, 1).value = re.sub(
                                r"\d{1,2}月\d{1,2}日\s*[-~至—]\s*\d{1,2}月\d{1,2}日",
                                self.week_title,
                                val,
                            )
            for slot in self.slots.values():
                if slot.current_holders != slot.original_holders:
                    if slot.sheet_name in wb.sheetnames:
                        ws = wb[slot.sheet_name]
                        ws[slot.cell_ref] = join_names(slot.current_holders, slot.original_raw_names)
            wb.save(output_path)
            return True, output_path
        except ImportError:
            return False, "缺少 openpyxl 库，无法直接写入 .xlsx 文件（状态已保存在本地 JSON 中）"
        except Exception as e:
            return False, f"导出 Excel 失败: {e}"
