# -*- coding: utf-8 -*-
"""
微信群替班消息自然语言解析与多班次判定引擎 (message_parser.py)
根据真实群聊语料全面增强：
1. 支持单条消息包含多个班次（换行分隔、逗号/顿号分隔），逐一分割解析并分别判定可接与不可接
2. 支持连班自动拆分与部分接取（如“今晚6-11服务台求替”拆为 18:00-20:00 + 20:00-23:00，若只满足一节则只接该节）
3. 支持全部时间格式：
   - 标准 24 小时制（15:30-18:00、20-23）
   - 12 小时制简写（8-10、8-11、6-8、6-10、6-11、2-4、2:00-3:30）
   - 混合制与缺省分钟（9-10:30、2-3:30、12:30-2:00）
   - 4位无冒号格式（1400-1530 1530 1800）
   - 中文口语（两点到六点、十点半到十二点半、九点到十点半）
4. 支持“二楼服务台 -> 服务台”、“（原王五班）-> 提取原班主”、跨行共享求替词/岗位/日期传播
5. 严格防误触：自动忽略自己发出的消息、引用回复消息、拍一拍/撤回等系统提示、纯换班消息及他人抢答（“我可以”）
6. 生成精准的多班部分接取回复文案（如全部可接回“我可以”，部分可接回“我可以 周一20:00-22:00、周三18:00-22:00”）
"""

from __future__ import annotations
import datetime
import json
import os
import re
import urllib.request
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set, Tuple

from schedule_manager import (
    ScheduleManager,
    ShiftSlot,
    WEEKDAYS,
    WEEKDAY_MAP,
)

CN_NUM_MAP = {
    "零": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5,
    "六": 6, "七": 7, "八": 8, "九": 9, "十": 10, "十一": 11, "十二": 12,
}


def parse_cn_or_int(token: str) -> Optional[int]:
    token = token.strip()
    if token.isdigit():
        return int(token)
    if token in CN_NUM_MAP:
        return CN_NUM_MAP[token]
    if token.startswith("十") and len(token) == 2 and token[1] in CN_NUM_MAP:
        return 10 + CN_NUM_MAP[token[1]]
    if token.endswith("十") and len(token) == 2 and token[0] in CN_NUM_MAP:
        return CN_NUM_MAP[token[0]] * 10
    if len(token) == 3 and token[1] == "十" and token[0] in CN_NUM_MAP and token[2] in CN_NUM_MAP:
        return CN_NUM_MAP[token[0]] * 10 + CN_NUM_MAP[token[2]]
    return None


@dataclass
class ShiftItemEvaluation:
    """单条消息分割出的单个班次（或单个连班子节）的解析与筛选结果"""
    segment_text: str
    weekday: str = ""
    date_str: str = ""
    post: str = ""
    start_min: int = 0
    end_min: int = 0
    parsed_time_str: str = ""
    detected_requester: str = ""
    matched_slots: List[ShiftSlot] = field(default_factory=list)
    should_take: bool = False
    reason: str = ""
    is_next_week: bool = False


@dataclass
class ParsedShiftRequest:
    """整条微信消息的汇总解析与判定结果（支持一条消息包含 1~N 个班次）"""
    raw_message: str
    sender_name: str = ""
    is_shift_request: bool = False
    reject_reason: str = ""

    # 所有被分割识别出的子班次判定列表
    items: List[ShiftItemEvaluation] = field(default_factory=list)
    accepted_items: List[ShiftItemEvaluation] = field(default_factory=list)
    rejected_items: List[ShiftItemEvaluation] = field(default_factory=list)

    # 汇总结果
    should_take: bool = False
    suggested_reply: str = ""          # 建议发送给微信群的精确回复文本（如 "我可以" 或 "我可以 周一20:00-22:00、周三18:00-22:00"）
    decision_reason: str = ""

    # 兼容单班展示属性
    @property
    def matched_slots(self) -> List[ShiftSlot]:
        out: List[ShiftSlot] = []
        seen: Set[str] = set()
        for it in self.accepted_items:
            for s in it.matched_slots:
                if s.slot_id not in seen:
                    seen.add(s.slot_id)
                    out.append(s)
        return out

    @property
    def detected_requester(self) -> str:
        for it in self.accepted_items:
            if it.detected_requester:
                return it.detected_requester
        for it in self.items:
            if it.detected_requester:
                return it.detected_requester
        return ""

    @property
    def weekday(self) -> str:
        wds = list(dict.fromkeys(it.weekday for it in self.items if it.weekday))
        return "、".join(wds)

    @property
    def date_str(self) -> str:
        ds = list(dict.fromkeys(it.date_str for it in self.items if it.date_str))
        return "、".join(ds)

    @property
    def post(self) -> str:
        ps = list(dict.fromkeys(it.post for it in self.items if it.post))
        return "、".join(ps)

    @property
    def parsed_time_str(self) -> str:
        ts = list(dict.fromkeys(it.parsed_time_str for it in self.items if it.parsed_time_str))
        return "、".join(ts)


class ShiftMessageParser:
    """替班消息自然语言解析器（支持多班分割、12/24小时制与真实微信群语料）"""

    def __init__(self, schedule_mgr: ScheduleManager, rules_path: str = "parser_rules.json"):
        self.schedule_mgr = schedule_mgr
        self.rules_path = os.path.abspath(rules_path)

        # 触发词与忽略词
        self.request_keywords: List[str] = [
            "替班", "求替", "代班", "求代", "替一下", "帮替", "有人替", "能替",
            "可以替", "有空吗", "有空的", "顶班", "顶一下", "出班", "求帮", "谁有空", "有人能", "可替",
        ]
        self.ignore_keywords: List[str] = [
            "我可以", "我来", "我接", "已替", "已找到", "已经找到", "不用了",
            "已解决", "谢谢", "收到", "好的", "替了", "我替", "先看到", "这边看到", "没问题",
        ]
        self.system_noise_patterns: List[str] = [
            r"撤回了一条消息",
            r"recalled a message",
            r"拍了拍",
            r"tickled",
            r"以下为新消息",
            r"查看更多消息",
            r"邀请.*加入了群聊",
            r"通过扫描.*加入群聊",
        ]
        self.ignore_swap_only: bool = True
        self.swap_keywords: List[str] = ["换班", "互换", "调班"]
        self.default_today_if_no_day: bool = True
        self.allow_partial_combo: bool = True      # 单条连班（如 6-11）只满足部分时段时，允许仅接满足的那一节
        self.partial_reply_with_time: bool = True  # 多班只接部分时，在回复中带上具体时间段以免歧义
        self.take_only_one_in_multi: bool = True   # 一条消息含多个可接班次时，仅接第1个，防止连抢被当脚本风控
        self.my_wechat_aliases: List[str] = []     # 用户额外的微信昵称/群昵称（用于防接自己发的消息）

        # 可选：大模型 API 语义兜底（仅当本地万能规则引擎无法识别极复杂口语时才调用，默认留空不启用）
        self.llm_api_url: str = ""                 # 例如 https://api.deepseek.com/v1/chat/completions 或 http://localhost:11434/v1/chat/completions
        self.llm_api_key: str = ""
        self.llm_model: str = ""

        self.load_rules()

    def load_rules(self) -> None:
        if not os.path.exists(self.rules_path):
            return
        try:
            with open(self.rules_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if "request_keywords" in data:
                self.request_keywords = list(data["request_keywords"])
            if "ignore_keywords" in data:
                self.ignore_keywords = list(data["ignore_keywords"])
            if "ignore_swap_only" in data:
                self.ignore_swap_only = bool(data["ignore_swap_only"])
            if "swap_keywords" in data:
                self.swap_keywords = list(data["swap_keywords"])
            if "default_today_if_no_day" in data:
                self.default_today_if_no_day = bool(data["default_today_if_no_day"])
            if "allow_partial_combo" in data:
                self.allow_partial_combo = bool(data["allow_partial_combo"])
            if "partial_reply_with_time" in data:
                self.partial_reply_with_time = bool(data["partial_reply_with_time"])
            if "take_only_one_in_multi" in data:
                self.take_only_one_in_multi = bool(data["take_only_one_in_multi"])
            if "my_wechat_aliases" in data:
                self.my_wechat_aliases = list(data["my_wechat_aliases"])
            self.llm_api_url = str(data.get("llm_api_url", "") or "").strip()
            self.llm_api_key = str(data.get("llm_api_key", "") or "").strip()
            self.llm_model = str(data.get("llm_model", "") or "").strip()
        except Exception as e:
            print(f"[警告] 读取解析规则失败: {e}")

    def save_rules(self) -> None:
        data = {
            "request_keywords": self.request_keywords,
            "ignore_keywords": self.ignore_keywords,
            "ignore_swap_only": self.ignore_swap_only,
            "swap_keywords": self.swap_keywords,
            "default_today_if_no_day": self.default_today_if_no_day,
            "allow_partial_combo": self.allow_partial_combo,
            "partial_reply_with_time": self.partial_reply_with_time,
            "take_only_one_in_multi": self.take_only_one_in_multi,
            "my_wechat_aliases": self.my_wechat_aliases,
            "llm_api_url": self.llm_api_url,
            "llm_api_key": self.llm_api_key,
            "llm_model": self.llm_model,
        }
        with open(self.rules_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def is_sender_self(self, sender_name: str, is_self_hint: bool = False) -> bool:
        """双重判断发送者是否为本人：UI气泡位置判断(is_self_hint) 或 姓名/微信昵称别名匹配"""
        if is_self_hint:
            return True
        s = sender_name.strip()
        if not s:
            return False
        op_name = self.schedule_mgr.operator_name.strip()
        if op_name and (s == op_name or op_name in s):
            return True
        for alias in self.my_wechat_aliases:
            a = alias.strip()
            if a and (s == a or a in s):
                return True
        return False

    # ------------------------------------------------------------------
    # 核心解析与多班评估主入口
    # ------------------------------------------------------------------
    def parse_and_evaluate(
        self,
        message: str,
        sender_name: str = "",
        now: Optional[datetime.datetime] = None,
        is_self_hint: bool = False,
        base_reply_word: str = "我可以",
    ) -> ParsedShiftRequest:
        if now is None:
            now = datetime.datetime.now()

        raw_msg = (message or "").strip()
        res = ParsedShiftRequest(raw_message=raw_msg, sender_name=sender_name.strip())

        if not raw_msg:
            res.reject_reason = "空消息"
            return res

        # 1. 严防自己接自己的消息（UI右对齐气泡 / 操作者姓名 / 微信别名）
        if self.is_sender_self(res.sender_name, is_self_hint=is_self_hint):
            res.reject_reason = "本人发送的消息，绝不接自己的班"
            return res

        # 2. 过滤微信系统提示（拍一拍 / 撤回消息 / 加入群聊等）
        for pat in self.system_noise_patterns:
            if re.search(pat, raw_msg, re.IGNORECASE):
                res.reject_reason = "微信系统通知消息，忽略"
                return res

        # 3. 处理微信“引用回复”与抢答/确认消息（防止引用他人求替消息时误触发）
        lines = [ln.strip() for ln in raw_msg.splitlines() if ln.strip()]
        first_line = lines[0] if lines else raw_msg
        # 如果第一行是简短抢答词或确认词（例如 "先看到这位同学的\n张三: 我可以" 或 "我可以\n引用..."）
        if any(ig in first_line for ig in self.ignore_keywords) and not any(rk in first_line for rk in self.request_keywords):
            res.reject_reason = f"命中忽略/确认词「{first_line[:20]}」"
            return res
        if len(raw_msg) <= 4 and any(k in raw_msg for k in ["1", "我", "来", "行", "ok", "OK"]):
            res.reject_reason = "简短抢答或闲聊回复"
            return res

        # 若微信引用消息把“引用原文”拼在了末尾，识别并截掉“引用 xxx 的消息：”之后的内容
        clean_msg = self._strip_quoted_tail(raw_msg)
        # 修复微信气泡折行/OCR跨行造成的断词（例如“周\n四”->“周四”，“2-\n330”->“2-330”）
        clean_msg = self._repair_wrapped_tokens(clean_msg)

        # 4. 忽略纯换班消息
        if self.ignore_swap_only and any(sk in clean_msg for sk in self.swap_keywords):
            if not any(rk in clean_msg for rk in ["替班", "求替", "替或换", "换或替"]):
                res.reject_reason = "疑似换班/调班请求（当前设置为仅接纯替班）"
                return res

        # 5. 提取消息全局上下文（原班主、发信人、全局岗位、全局日期、求替意图）
        global_names = self._extract_all_assistant_names(clean_msg, res.sender_name)
        global_post = self._extract_post(clean_msg)
        global_wd, global_date, global_next_week = self._extract_weekday(clean_msg, now)
        has_req_kw = any(rk in clean_msg for rk in self.request_keywords)

        # 6. 分割为独立班次片段（支持多行 + 逗号/顿号/分号，且剥离括号里的拆分注释如 （1400-1530 1530 1800））
        segments = self._split_into_shift_segments(clean_msg)
        if not segments:
            segments = [clean_msg]

        # 从所有片段中提取班次要素（支持“周一、周三 18-20”和“周三 18-20、20-22”这种共享星期/时间传播）
        extracted_candidates: List[ShiftItemEvaluation] = []
        pending_weekdays: List[Tuple[str, str, bool]] = []
        last_wd, last_date, last_next_week = global_wd, global_date, global_next_week

        for seg in segments:
            seg_clean, bracket_names, bracket_sub_times = self._extract_and_clean_brackets(seg)
            seg_wds = self._extract_all_weekdays(seg_clean, now)
            seg_post = self._extract_post(seg_clean)
            time_ranges = self._extract_all_time_ranges(seg_clean, seg_post or global_post)

            if not time_ranges and bracket_sub_times:
                time_ranges = bracket_sub_times

            # 如果本段只有星期/日期没有时间（例如 "周一、周三 18-20 四楼求替" 中的 "周一"）
            if not time_ranges:
                if seg_wds:
                    pending_weekdays.extend(seg_wds)
                    last_wd, last_date, last_next_week = seg_wds[-1]
                continue

            if seg_wds:
                target_days = pending_weekdays + seg_wds
                pending_weekdays = []
                last_wd, last_date, last_next_week = seg_wds[-1]
            elif pending_weekdays:
                target_days = pending_weekdays
                pending_weekdays = []
            elif last_wd:
                target_days = [(last_wd, last_date, last_next_week)]
            elif self.default_today_if_no_day:
                today_wd = WEEKDAYS[now.weekday()]
                target_days = [(today_wd, self.schedule_mgr.weekday_dates.get(today_wd, ("今天", ""))[0], False)]
            else:
                target_days = [("", "", False)]

            use_post = seg_post or global_post
            seg_names = self._extract_all_assistant_names(seg_clean, "")
            req_candidates = list(dict.fromkeys(bracket_names + seg_names + global_names))

            for use_wd, use_date, use_next_week in target_days:
                for s_min, e_min, t_str in time_ranges:
                    resolved_date = self._resolve_target_date(use_wd, use_date, use_next_week)
                    ev = ShiftItemEvaluation(
                        segment_text=seg,
                        weekday=use_wd,
                        date_str=resolved_date,
                        post=use_post,
                        start_min=s_min,
                        end_min=e_min,
                        parsed_time_str=t_str,
                        detected_requester=req_candidates[0] if req_candidates else "",
                        is_next_week=use_next_week,
                    )
                    if not use_wd:
                        ev.should_take = False
                        ev.reason = "未识别出具体星期或日期"
                        extracted_candidates.append(ev)
                        continue

                    self._evaluate_single_shift_item(ev, req_candidates, now, extracted_candidates)

        # 若有求替班关键词，但未提取到明确数字时间，尝试【排班表模糊时段反查】与【可选AI语义兜底】
        if not extracted_candidates and has_req_kw:
            self._try_schedule_fuzzy_lookup(
                clean_msg, global_wd, global_date, global_next_week, global_post, global_names, now, extracted_candidates
            )
        if not extracted_candidates and has_req_kw and self.llm_api_url and self.llm_model:
            self._try_llm_extract(clean_msg, res.sender_name, now, global_names, extracted_candidates)

        # 判断是否构成替班请求
        if not extracted_candidates:
            if has_req_kw:
                res.is_shift_request = True
                res.should_take = False
                res.decision_reason = "检测到求替班意图，但未能从消息中识别出有效时间段或班次"
            else:
                res.is_shift_request = False
                res.reject_reason = "未检测到替班意图或班次要素"
            return res

        # 若无求替关键词，要求至少有一个班次同时具备【星期 + 时间 + 岗位】（防止普通聊天报时误触发）
        if not has_req_kw:
            has_complete = any(it.weekday and it.parsed_time_str and it.post for it in extracted_candidates)
            if not has_complete:
                res.is_shift_request = False
                res.reject_reason = "未检测到求替班关键词且缺少完整班次要素"
                return res

        # 7. 去重（同一条消息内若出现完全重叠的 slot 组合，避免重复计数）
        res.is_shift_request = True
        res.items = self._deduplicate_evaluations(extracted_candidates)

        # 多班消息防脚本风控：一条消息若拆出多个可接班次，默认只接第1个，其余自动跳过
        if self.take_only_one_in_multi:
            ok_items = [it for it in res.items if it.should_take]
            if len(ok_items) > 1:
                for extra_it in ok_items[1:]:
                    extra_it.should_take = False
                    extra_it.reason = "多班消息仅接1个班（防连抢被当脚本，其余自动跳过）"

        res.accepted_items = [it for it in res.items if it.should_take]
        res.rejected_items = [it for it in res.items if not it.should_take]

        # 8. 生成最终判定与精准回复文本
        if not res.accepted_items:
            res.should_take = False
            reasons = "；".join(
                f"[{it.weekday} {it.parsed_time_str} {it.post or '未知岗位'}] {it.reason}"
                for it in res.rejected_items
            )
            res.decision_reason = f"不接此班：{reasons}"
            return res

        res.should_take = True
        res.suggested_reply = self._build_reply_message(res, base_reply_word)

        acc_desc = "、".join(
            f"{it.weekday}({it.date_str}) {it.post} {it.parsed_time_str}"
            for it in res.accepted_items
        )
        if not res.rejected_items:
            res.decision_reason = f"✅ 全部符合条件！可接：{acc_desc} -> 回复「{res.suggested_reply}」"
        else:
            rej_desc = "；".join(
                f"{it.weekday} {it.parsed_time_str} {it.post}({it.reason})"
                for it in res.rejected_items
            )
            res.decision_reason = (
                f"✅ 部分班次符合条件！\n"
                f"   可接：{acc_desc}\n"
                f"   跳过：{rej_desc}\n"
                f"   自动回复：「{res.suggested_reply}」"
            )
        return res

    # ------------------------------------------------------------------
    # 单个班次 / 连班评估（含连班拆分接取逻辑）
    # ------------------------------------------------------------------
    def _evaluate_single_shift_item(
        self,
        ev: ShiftItemEvaluation,
        req_candidates: List[str],
        now: datetime.datetime,
        out_list: List[ShiftItemEvaluation],
    ) -> None:
        """在排班表中定位匹配的槽位并判定可接性；若遇连班部分可接且开启拆分，自动拆成子班次！"""
        # 依次尝试候选姓名定位（比如先查括号里的“原王五班”，再查发信人“赵六”，最后不限人名按时间+岗位查）
        combos: List[List[ShiftSlot]] = []
        matched_requester = ev.detected_requester

        for name in req_candidates:
            combos = self.schedule_mgr.find_matching_slots_for_request(
                weekday=ev.weekday,
                start_min=ev.start_min,
                end_min=ev.end_min,
                raw_time_token=ev.parsed_time_str,
                post=ev.post if ev.post else None,
                requester_name=name,
            )
            if combos:
                matched_requester = name
                break

        if not combos:
            combos = self.schedule_mgr.find_matching_slots_for_request(
                weekday=ev.weekday,
                start_min=ev.start_min,
                end_min=ev.end_min,
                raw_time_token=ev.parsed_time_str,
                post=ev.post if ev.post else None,
                requester_name=None,
            )

        if not combos:
            ev.should_take = False
            ev.reason = f"排班表中未找到匹配班次 ({ev.weekday} {ev.parsed_time_str} {ev.post or '未知岗位'})"
            out_list.append(ev)
            return

        # 若消息未写岗位且发信人无法唯一确定岗位（同时命中多个不同楼层）
        if not ev.post and len({c[0].post for c in combos}) > 1:
            possible_posts = sorted({c[0].post for c in combos})
            ev.matched_slots = combos[0]
            ev.post = f"未注明({'/'.join(possible_posts)})"
            ev.should_take = False
            ev.reason = f"未写明楼层/岗位且无法由姓名唯一定位（候选：{'、'.join(possible_posts)}），为防搞混暂不接"
            out_list.append(ev)
            return

        combo = combos[0]
        if not ev.post:
            ev.post = combo[0].post

        # 如果当前只有一位值班人（或候选人之一在 current_holders 中），精确锁定被替班人姓名
        if matched_requester and any(matched_requester in s.current_holders for s in combo):
            ev.detected_requester = matched_requester
        elif len(combo[0].current_holders) == 1:
            ev.detected_requester = combo[0].current_holders[0]
        elif req_candidates:
            for c_name in req_candidates:
                if any(c_name in s.original_holders or c_name in s.current_holders for s in combo):
                    ev.detected_requester = c_name
                    break

        # 逐个子节检验状态（按每周周期排班逻辑，不再以“时间已过”屏蔽，以兼容下周预换班）
        statuses = [
            (s, *self.schedule_mgr.evaluate_slot_status(s, check_expired=False, now=now))
            for s in combo
        ]

        # 情况1：整段（单班或完整连班）全部可接！
        if all(st == "AVAILABLE" for _, st, _ in statuses):
            ev.matched_slots = combo
            ev.should_take = True
            ev.reason = "符合条件，可接班"
            out_list.append(ev)
            return

        # 情况2：是连班（len > 1），全部不可接
        if not any(st == "AVAILABLE" for _, st, _ in statuses):
            first_bad = next((s, st, r) for s, st, r in statuses if st != "AVAILABLE")
            ev.matched_slots = combo
            ev.should_take = False
            ev.reason = first_bad[2]
            out_list.append(ev)
            return

        # 情况3：连班中部分子班次可接、部分子班次不可接（如图15：今晚6-11服务台求替，只能接6-8）
        if self.allow_partial_combo and len(combo) > 1:
            for s, st, r in statuses:
                sub_ev = ShiftItemEvaluation(
                    segment_text=ev.segment_text,
                    weekday=s.weekday,
                    date_str=ev.date_str or s.date_str,
                    post=s.post,
                    start_min=s.start_min,
                    end_min=s.end_min,
                    parsed_time_str=s.time_str,
                    detected_requester=ev.detected_requester if ev.detected_requester in s.current_holders else (
                        s.current_holders[0] if len(s.current_holders) == 1 else ev.detected_requester
                    ),
                    matched_slots=[s],
                    should_take=(st == "AVAILABLE"),
                    reason="符合条件（连班拆分接取）" if st == "AVAILABLE" else r,
                    is_next_week=ev.is_next_week,
                )
                out_list.append(sub_ev)
        else:
            first_bad = next((s, st, r) for s, st, r in statuses if st != "AVAILABLE")
            ev.matched_slots = combo
            ev.should_take = False
            ev.reason = f"连班中存在不可接时段({first_bad[0].time_str}: {first_bad[2]})"
            out_list.append(ev)

    def _deduplicate_evaluations(self, items: List[ShiftItemEvaluation]) -> List[ShiftItemEvaluation]:
        """按 matched_slots 的 slot_id 集合去重"""
        out: List[ShiftItemEvaluation] = []
        seen_keys: Set[Tuple[str, ...]] = set()
        for it in items:
            if it.matched_slots:
                key = tuple(s.slot_id for s in it.matched_slots)
            else:
                key = (it.weekday, it.parsed_time_str, it.post, it.reason)
            if key not in seen_keys:
                seen_keys.add(key)
                out.append(it)
        return out

    def _format_casual_12h_time(self, it: ShiftItemEvaluation) -> str:
        """
        将班次时段格式化为同学群聊最自然口语的 12 小时制简写：
        - 整点只写 1~2 位数字（如 9、12、2、6、8），半点加 30（如 1030、330、1230、930）
        - 示例：09:00-12:00 -> 9-12，14:00-15:30 -> 2-330，18:00-20:00 -> 6-8，20:00-23:00 -> 8-11
        - D5 图书室保留“上午/晚上”
        """
        if it.start_min < 0 or it.end_min <= 0:
            return it.parsed_time_str

        def _fmt_pt(total_min: int) -> str:
            h24, m = divmod(total_min, 60)
            h12 = h24 if h24 <= 12 else h24 - 12
            if h12 == 0:
                h12 = 12
            return str(h12) if m == 0 else f"{h12}{m:02d}"

        return f"{_fmt_pt(it.start_min)}-{_fmt_pt(it.end_min)}"

    def _build_reply_message(self, res: ParsedShiftRequest, base_reply_word: str) -> str:
        """
        根据接班结果生成最自然、无歧义且防脚本风控的微信回复文本：
        - 单个求替班次全部接下：直接发 "我可以"
        - 原消息含多个班次或连班拆分，只接其中1个/部分：用口语12小时制附上具体星期与简写时间，
          例如 "我可以 周日9-12"、"我可以 周四2-330"、"我可以 6-8"
        """
        word = (base_reply_word or "我可以").strip()
        if len(res.items) == 1 and not res.rejected_items:
            return word

        parts: List[str] = []
        all_same_day = len({(it.weekday, it.is_next_week) for it in res.items}) == 1
        for it in res.accepted_items:
            t_casual = self._format_casual_12h_time(it)
            wd_label = f"下{it.weekday}" if it.is_next_week else it.weekday
            if all_same_day:
                parts.append(t_casual if self.partial_reply_with_time else wd_label)
            elif self.partial_reply_with_time:
                parts.append(f"{wd_label}{t_casual}")
            else:
                parts.append(wd_label)

        parts = list(dict.fromkeys(parts))
        return f"{word} {'、'.join(parts)}"

    # ------------------------------------------------------------------
    # 文本预处理、分段与多格式时间解析工具
    # ------------------------------------------------------------------
    def _resolve_target_date(self, wd: str, explicit_date: str, is_next_week: bool) -> str:
        info = self.schedule_mgr.weekday_dates.get(wd, ("", ""))
        this_week_date, this_week_iso = info[0], info[1]
        if not is_next_week:
            return explicit_date or this_week_date
        # 若 explicit_date 只是默认填入的本周日期（如用户只说了“下周二”），应推算 +7 天为下周真实日期
        if explicit_date and explicit_date != this_week_date:
            return f"{explicit_date}(下周)" if "(下周)" not in explicit_date else explicit_date
        if this_week_iso:
            try:
                dt = datetime.date.fromisoformat(this_week_iso) + datetime.timedelta(days=7)
                return f"{dt.month}月{dt.day}日(下周)"
            except Exception:
                pass
        return f"{this_week_date}(下周)" if this_week_date else "下周"

    def _repair_wrapped_tokens(self, msg: str) -> str:
        """合并微信自动折行 / OCR 跨行造成的断词（如 '周\\n四'、'2-\\n330'、'四楼\\n书库'）"""
        s = re.sub(r"周\s*\n\s*(?=[一二三四五六日天末1-7])", "周", msg)
        s = re.sub(r"(星期|礼拜)\s*\n\s*(?=[一二三四五六日天])", r"\1", s)
        s = re.sub(r"([上下早中晚今明后大])\s*\n\s*(?=[午上早晚天])", r"\1", s)
        s = re.sub(r"([一二两三四五六七八九十\d])\s*\n\s*(?=[楼层])", r"\1", s)
        s = re.sub(r"(服)\s*\n\s*(?=务)|(服务)\s*\n\s*(?=台)|(书)\s*\n\s*(?=库)|(图)\s*\n\s*(?=书)|(图书)\s*\n\s*(?=室)", lambda m: m.group(0).replace("\n", "").strip(), s)
        s = re.sub(r"([求代帮顶出可])\s*\n\s*(?=[替班代])", r"\1", s)
        s = re.sub(r"(\d)\s*\n\s*(?=[-~～—–−至到:：\.])|([-~～—–−至到:：])\s*\n\s*(?=\d)", lambda m: re.sub(r"\s*\n\s*", "", m.group(0)), s)
        return s

    def _strip_quoted_tail(self, msg: str) -> str:
        """如果微信消息末尾拼了引用分隔，尝试去除引用尾部"""
        # 常见的微信 UIA 引用拼接格式：正文\n引用 XXX 的消息：... 或 正文\n| XXX: ...
        m = re.split(r"\n\s*(?:引用\s+.*|[\|｜]\s*[^:\n]{1,20}:)", msg, maxsplit=1)
        return m[0].strip() if m else msg

    def _split_into_shift_segments(self, msg: str) -> List[str]:
        """
        把一条可能包含多个班次的消息切分成独立片段（多行、逗号/顿号/分号，以及无标点直连的连续星期项）。
        注意保护括号内部（如 （1400-1530 1530 1800）），不从括号内部切开！
        """
        msg = self._repair_wrapped_tokens(msg)
        placeholders: List[str] = []

        def _mask(m):
            placeholders.append(m.group(0))
            return f"__BRK_{len(placeholders)-1}__"

        masked = re.sub(r"[（\(][^（）\(\)]*[）\)]", _mask, msg)
        raw_parts = re.split(r"[\n\r,，;；、]+", masked)

        day_start_re = re.compile(
            r"(?<=[\d楼库台室午晚班替])(?=(?:下*(?:周|星期|礼拜)[一二三四五六日天]|今天|今日|今早|今晚|明天|明日|明早|明晚|后天|后晚))"
        )

        segments: List[str] = []
        for p in raw_parts:
            p = p.strip()
            if not p:
                continue
            for sub in day_start_re.split(p):
                sub = sub.strip()
                if not sub:
                    continue
                for idx, orig in enumerate(placeholders):
                    sub = sub.replace(f"__BRK_{idx}__", orig)
                segments.append(sub)
        return segments

    def _extract_and_clean_brackets(
        self, seg: str
    ) -> Tuple[str, List[str], List[Tuple[int, int, str]]]:
        """
        提取括号内信息：
        - 提取括号内的助理姓名（如 （原王五班） -> ['王五']）
        - 提取括号内的 4 位时间拆解（如 （1400-1530 1530 1800） -> 备选子时间）
        - 返回剥离括号后的主文本（防止与括号外的“两点到六点”重复识别！）
        """
        bracket_names: List[str] = []
        bracket_times: List[Tuple[int, int, str]] = []
        sorted_names = sorted(self.schedule_mgr.all_names, key=len, reverse=True)

        def _proc(m):
            inner = m.group(1).strip()
            for nm in sorted_names:
                if nm in inner and nm not in bracket_names:
                    bracket_names.append(nm)
            # 若括号内包含时间格式（且不含姓名），记录为备用时间并从主句剥离
            ts = self._extract_all_time_ranges(inner, "")
            if ts:
                bracket_times.extend(ts)
            return " "

        cleaned = re.sub(r"[（\(]([^（）\(\)]*)[）\)]", _proc, seg)
        return cleaned.strip(), bracket_names, bracket_times

    def _extract_all_assistant_names(self, msg: str, sender: str) -> List[str]:
        """提取所有匹配排班表的助理姓名（按出现顺序与长度优先）"""
        found: List[str] = []
        sorted_names = sorted(self.schedule_mgr.all_names, key=len, reverse=True)
        for name in sorted_names:
            if sender and name in sender and name not in found:
                found.append(name)
        for name in sorted_names:
            if name in msg and name not in found:
                found.append(name)
        return found

    def _extract_post(self, msg: str) -> str:
        """岗位提取：只要带“服务台/前台/总台”优先识别为服务台（解决图1“二楼服务台”的歧义！）"""
        if any(k in msg for k in ["服务台", "前台", "总台", "大厅"]):
            return "服务台"
        if any(k in msg for k in ["2楼", "二楼", "2F", "2f", "二库"]):
            return "2楼书库"
        if any(k in msg for k in ["3楼", "三楼", "3F", "3f", "三库"]):
            return "3楼书库"
        if any(k in msg for k in ["4楼", "四楼", "4F", "4f", "四库"]):
            return "4楼书库"
        if any(k in msg.upper() for k in ["D5", "图书室"]):
            return "D5图书室"
        return ""

    def _extract_all_weekdays(self, msg: str, now: datetime.datetime) -> List[Tuple[str, str, bool]]:
        """提取片段中提及的所有星期/日期（例如 '周三周四下午2-4'）"""
        is_next_week = "下周" in msg or "下星期" in msg or "下礼拜" in msg
        found: List[Tuple[str, str, bool]] = []
        seen_wd: Set[str] = set()

        for m_wd in re.finditer(r"(周[一二三四五六日天123456]|星期[一二三四五六日天]|礼拜[一二三四五六日天123456])", msg):
            wd_token = m_wd.group(1)
            wd = WEEKDAYS[WEEKDAY_MAP[wd_token]]
            if wd not in seen_wd:
                seen_wd.add(wd)
                found.append((wd, self.schedule_mgr.weekday_dates.get(wd, ("", ""))[0], is_next_week))

        if found:
            return found

        wd, d_str, nw = self._extract_weekday(msg, now)
        if wd:
            return [(wd, d_str, nw)]
        return []

    def _extract_weekday(self, msg: str, now: datetime.datetime) -> Tuple[str, str, bool]:
        is_next_week = "下周" in msg or "下星期" in msg or "下礼拜" in msg

        m_wd = re.search(r"(周[一二三四五六日天123456]|星期[一二三四五六日天]|礼拜[一二三四五六日天123456])", msg)
        if m_wd:
            wd_token = m_wd.group(1)
            wd = WEEKDAYS[WEEKDAY_MAP[wd_token]]
            return wd, self.schedule_mgr.weekday_dates.get(wd, ("", ""))[0], is_next_week

        today_idx = now.weekday()
        if "大后天" in msg:
            idx = (today_idx + 3) % 7
            wd = WEEKDAYS[idx]
            return wd, self.schedule_mgr.weekday_dates.get(wd, ("", ""))[0], (today_idx + 3 >= 7)
        if any(k in msg for k in ["后天", "后日", "后晚"]):
            idx = (today_idx + 2) % 7
            wd = WEEKDAYS[idx]
            return wd, self.schedule_mgr.weekday_dates.get(wd, ("", ""))[0], (today_idx + 2 >= 7)
        if any(k in msg for k in ["明天", "明日", "明早", "明晚", "明天的"]):
            idx = (today_idx + 1) % 7
            wd = WEEKDAYS[idx]
            return wd, self.schedule_mgr.weekday_dates.get(wd, ("", ""))[0], (today_idx + 1 >= 7)
        if any(k in msg for k in ["今天", "今日", "今晚", "今早", "今午"]):
            wd = WEEKDAYS[today_idx]
            return wd, self.schedule_mgr.weekday_dates.get(wd, ("", ""))[0], False

        # 几月几日 / 11.18
        m_date = re.search(r"(\d{1,2})\s*[月\./]\s*(\d{1,2})\s*[日号]?", msg)
        if m_date:
            mo, da = int(m_date.group(1)), int(m_date.group(2))
            if 1 <= mo <= 12 and 1 <= da <= 31:
                target_str = f"{mo}月{da}日"
                for wd, (d_str, _) in self.schedule_mgr.weekday_dates.items():
                    if d_str == target_str:
                        return wd, d_str, False
                try:
                    dt = datetime.date(now.year, mo, da)
                    wd = WEEKDAYS[dt.weekday()]
                    return wd, target_str, False
                except ValueError:
                    pass

        m_day_only = re.search(r"(?<!\d)(\d{1,2})\s*[日号]", msg)
        if m_day_only:
            da = int(m_day_only.group(1))
            for wd, (d_str, _) in self.schedule_mgr.weekday_dates.items():
                if d_str.endswith(f"月{da}日"):
                    return wd, d_str, False

        return "", "", False

    def _try_schedule_fuzzy_lookup(
        self,
        clean_msg: str,
        global_wd: str,
        global_date: str,
        global_next_week: bool,
        global_post: str,
        global_names: List[str],
        now: datetime.datetime,
        out_list: List[ShiftItemEvaluation],
    ) -> None:
        """
        兜底方案1（无需训练数据，直接基于当前排班表反查）：
        当消息里未写出明确数字时间，但包含“早班/早八/午班/闭馆班”或“发信人+星期几+上下午/晚上”时，
        直接在加载好的排班表 slots 中反查唯一匹配的班次。
        """
        weekday_names = ("周一", "周二", "周三", "周四", "周五", "周六", "周日")
        target_wd = global_wd
        target_date = self._resolve_target_date(target_wd, global_date, global_next_week) if target_wd else global_date
        if not target_wd:
            # 只有明确含今天语义词或特定班次代称（早班/午班/闭馆班/今晚/今早）时才默认今天
            if any(k in clean_msg for k in ["今天", "今日", "今晚", "今早", "早班", "早八", "午班", "闭馆"]):
                target_wd = weekday_names[now.weekday()]
                target_date = f"{now.month}月{now.day}日"
            else:
                return

        # 识别班次时段代称或模糊时段词
        period_kind = ""
        if any(k in clean_msg for k in ["早班", "早八", "早8", "第一班"]):
            period_kind = "early"
        elif any(k in clean_msg for k in ["早九", "早9"]):
            period_kind = "morning9"
        elif any(k in clean_msg for k in ["午班", "中午班", "中午"]):
            period_kind = "noon"
        elif any(k in clean_msg for k in ["闭馆", "关门", "晚十", "最晚"]):
            period_kind = "closing"
        elif any(k in clean_msg for k in ["晚上", "晚班", "今晚", "明晚", "后晚", "夜班"]):
            period_kind = "evening"
        elif any(k in clean_msg for k in ["下午", "午后"]):
            period_kind = "afternoon"
        elif any(k in clean_msg for k in ["上午", "早上", "今早", "明早"]):
            period_kind = "morning"

        if not period_kind and not global_names:
            return

        # 在排班表中检索符合条件的所有槽位
        candidates: List[ShiftSlot] = []
        for s in self.schedule_mgr.slots.values():
            if s.weekday != target_wd:
                continue
            if global_post and s.post != global_post:
                continue
            if period_kind == "early":
                if s.start_min != 480:
                    continue
            elif period_kind == "morning9":
                if s.start_min != 540:
                    continue
            elif period_kind == "noon":
                if s.start_min not in (720, 750):
                    continue
            elif period_kind == "closing":
                if s.end_min != 1380:
                    continue
            elif period_kind == "morning":
                if not ((0 <= s.start_min < 720 and s.end_min <= 780) or s.start_min == -1):
                    continue
            elif period_kind == "afternoon":
                if not (720 <= s.start_min < 1080):
                    continue
            elif period_kind == "evening":
                if not (s.start_min >= 1080 or s.start_min == -2):
                    continue
            candidates.append(s)

        if not candidates:
            return

        # 优先按发信人/原值班人筛选，若无法按人名筛选且是“早班/午班/闭馆班”且同岗位同时间唯一也可直接锁定
        req_matched: List[ShiftSlot] = []
        if global_names:
            for nm in global_names:
                hit = [s for s in candidates if nm in s.current_holders]
                if hit:
                    req_matched = hit
                    break

        chosen_pool = req_matched
        if not chosen_pool:
            if period_kind in ("early", "morning9", "noon", "closing"):
                unique_slots = {(s.post, s.start_min, s.end_min, s.time_str) for s in candidates}
                if len(unique_slots) == 1:
                    chosen_pool = candidates
            else:
                return

        if not chosen_pool:
            return

        posts_in_pool = {s.post for s in chosen_pool}
        by_key: Dict[Tuple[str, int, int, str], ShiftSlot] = {}
        for s in chosen_pool:
            by_key[(s.post, s.start_min, s.end_min, s.time_str)] = s

        ordered = sorted(by_key.values(), key=lambda x: (x.post, x.start_min, x.end_min))
        if len(ordered) > 1 and len(posts_in_pool) == 1 and ordered[0].start_min >= 0:
            is_chain = all(
                ordered[i].end_min == ordered[i + 1].start_min
                for i in range(len(ordered) - 1)
            )
            if is_chain:
                merged_st = ordered[0].start_min
                merged_ed = ordered[-1].end_min
                merged_str = f"{merged_st // 60:02d}:{merged_st % 60:02d}-{merged_ed // 60:02d}:{merged_ed % 60:02d}"
                ev = ShiftItemEvaluation(
                    segment_text=clean_msg,
                    weekday=target_wd,
                    date_str=target_date or self.schedule_mgr.weekday_dates.get(target_wd, ("", ""))[0],
                    post=global_post or ordered[0].post,
                    start_min=merged_st,
                    end_min=merged_ed,
                    parsed_time_str=merged_str,
                    detected_requester=global_names[0] if global_names else "",
                    is_next_week=global_next_week,
                )
                self._evaluate_single_shift_item(ev, global_names, now, out_list)
                return

        for s in ordered:
            ev = ShiftItemEvaluation(
                segment_text=clean_msg,
                weekday=s.weekday,
                date_str=target_date or s.date_str,
                post=global_post or s.post,
                start_min=s.start_min,
                end_min=s.end_min,
                parsed_time_str=s.time_str,
                detected_requester=global_names[0] if global_names else "",
                is_next_week=global_next_week,
            )
            self._evaluate_single_shift_item(ev, global_names, now, out_list)

    def _try_llm_extract(
        self,
        clean_msg: str,
        sender_name: str,
        now: datetime.datetime,
        req_candidates: List[str],
        out_list: List[ShiftItemEvaluation],
    ) -> None:
        """
        兜底方案2（可选 AI 语义解析）：
        仅当本地规则与排班表反查均未能识别出明确班次，且用户主动配置了 OpenAI 兼容 API 时调用，
        用大模型把任意非标口语抽成标准 JSON [{weekday, post, start, end, requester}]。
        """
        if not self.llm_api_url or not self.llm_model:
            return

        weekday_names = ("周一", "周二", "周三", "周四", "周五", "周六", "周日")
        today_wd = weekday_names[now.weekday()]
        today_str = f"{now.year}-{now.month:02d}-{now.day:02d} ({today_wd})"
        posts_hint = "服务台, 2楼书库, 3楼书库, 4楼书库, D5图书室"

        sys_prompt = (
            "你是图书馆排班群消息结构化解析助手。请把用户发出的求替班消息解析成 JSON 数组，不要输出多余文字或代码块标记。\n"
            f"当前时间：{today_str}。可选岗位：[{posts_hint}]。\n"
            "每个元素格式：{\"weekday\": \"周一~周日之一\", \"post\": \"岗位名称或空字符串\", "
            "\"start\": \"24小时制HH:MM 或 上午/晚上\", \"end\": \"24小时制HH:MM 或 空字符串\", \"requester\": \"原值班人姓名或空字符串\"}。\n"
            "注意：图书馆开放时间为 08:00-23:00，像 2-4、2-330 是 14:00-16:00、14:00-15:30；9-1030 是 09:00-10:30；"
            "若只写了D5上午或晚上，start填\"上午\"或\"晚上\"，end留空；若无法判断星期默认填今天；若非求替班返回 []。"
        )
        user_prompt = f"发信人：{sender_name or '未知'}\n消息内容：{clean_msg}"

        url = self.llm_api_url.strip().rstrip("/")
        if not url.endswith("/chat/completions"):
            url = f"{url}/chat/completions"

        payload = json.dumps(
            {
                "model": self.llm_model.strip(),
                "temperature": 0,
                "messages": [
                    {"role": "system", "content": sys_prompt},
                    {"role": "user", "content": user_prompt},
                ],
            },
            ensure_ascii=False,
        ).encode("utf-8")

        headers = {"Content-Type": "application/json"}
        if self.llm_api_key.strip():
            headers["Authorization"] = f"Bearer {self.llm_api_key.strip()}"

        try:
            req = urllib.request.Request(url, data=payload, headers=headers, method="POST")
            with urllib.request.urlopen(req, timeout=6) as resp:
                raw_body = resp.read().decode("utf-8", errors="ignore")
            data = json.loads(raw_body)
            content = data["choices"][0]["message"]["content"].strip()
            content = re.sub(r"^```(?:json)?\s*|\s*```$", "", content, flags=re.IGNORECASE | re.DOTALL).strip()
            parsed_items = json.loads(content)
            if not isinstance(parsed_items, list):
                return
        except Exception:
            return

        for item in parsed_items:
            if not isinstance(item, dict):
                continue
            wd = str(item.get("weekday") or today_wd).strip()
            if wd not in weekday_names:
                wd = today_wd
            post = str(item.get("post") or "").strip()
            st_raw = str(item.get("start") or "").strip()
            ed_raw = str(item.get("end") or "").strip()
            req_name = str(item.get("requester") or "").strip()

            start_min, end_min, t_str = 0, 0, ""
            if st_raw == "上午":
                start_min, end_min, t_str = -1, -1, "上午"
            elif st_raw == "晚上":
                start_min, end_min, t_str = -2, -2, "晚上"
            else:
                m1 = re.fullmatch(r"(\d{1,2}):(\d{2})", st_raw)
                m2 = re.fullmatch(r"(\d{1,2}):(\d{2})", ed_raw)
                if not (m1 and m2):
                    continue
                h1, min1 = int(m1.group(1)), int(m1.group(2))
                h2, min2 = int(m2.group(1)), int(m2.group(2))
                if not (0 <= h1 <= 24 and 0 <= min1 < 60 and 0 <= h2 <= 24 and 0 <= min2 < 60):
                    continue
                start_min = h1 * 60 + min1
                end_min = h2 * 60 + min2
                t_str = f"{h1:02d}:{min1:02d}-{h2:02d}:{min2:02d}"

            date_str = self.schedule_mgr.weekday_dates.get(wd, ("", ""))[0]
            local_cands = list(req_candidates)
            if req_name and req_name not in local_cands:
                local_cands.insert(0, req_name)
            ev = ShiftItemEvaluation(
                segment_text=clean_msg,
                weekday=wd,
                date_str=date_str,
                post=post,
                start_min=start_min,
                end_min=end_min,
                parsed_time_str=t_str,
                detected_requester=req_name or (local_cands[0] if local_cands else ""),
            )
            self._evaluate_single_shift_item(ev, local_cands, now, out_list)

    def _parse_single_time_token(self, text: str) -> List[dict]:
        """
        统一分词扫描消息中的单个时间点（中文点半 / 冒号 / 合法点号 / 3-4位紧缩数字 / 纯数字小时），
        同时记录位置区间、前缀“上午/下午/晚上”修饰及是否显式时间格式。
        """
        tokens: List[dict] = []
        token_regex = re.compile(
            r"(?:"
            # 分支1：中文或数字点钟，例如 9点半、十点半、十二点、两点、下午6点、10点30分
            r"(?:(?P<p1>上午|早上|早|中午|午|下午|晚上|晚)\s*)?(?<!\d)"
            r"(?P<d_h>二十[一二三四]?|十[一二]?|[一二两三四五六七八九]|\d{1,2})\s*[点时]"
            r"(?:\s*(?P<d_m>半|\d{1,2})\s*分?)?"
            r"|"
            # 分支2：带冒号或合法点号的时刻，例如 15:30、9：00、9.30、14.00
            r"(?:(?P<p2>上午|早上|早|中午|午|下午|晚上|晚)\s*)?(?<!\d)"
            r"(?P<c_h>\d{1,2})\s*(?P<c_sep>[:：\.])\s*(?P<c_m>\d{2})(?!\d)(?![月日号楼层F])"
            r"|"
            # 分支3：3位或4位紧缩数字时刻，例如 1030、930、1400、1530、0900
            r"(?:(?P<p3>上午|早上|早|中午|午|下午|晚上|晚)\s*)?(?<!\d)"
            r"(?P<m_tok>\d{3,4})(?!(?:\.\d|[:：\d月日号楼层F]))"
            r"|"
            # 分支4：纯 1~2 位数字小时，例如 9、10、18、2（排除周2、4楼、18号、11月）
            r"(?:(?P<p4>上午|早上|早|中午|午|下午|晚上|晚)\s*)?(?<![\d周])(?<!星期)(?<!礼拜)"
            r"(?P<b_h>\d{1,2})(?!(?:\.\d|[:：\d月日号楼层个位人名天周F]))"
            r")"
        )

        pm_words = ("下午", "晚上", "晚")
        am_words = ("上午", "早上", "早", "中午", "午")

        for m in token_regex.finditer(text):
            prefix = m.group("p1") or m.group("p2") or m.group("p3") or m.group("p4") or ""
            tok_pm = any(w in prefix for w in pm_words)
            tok_am = any(w in prefix for w in am_words)

            if m.group("d_h") is not None:
                h = parse_cn_or_int(m.group("d_h"))
                m_raw = m.group("d_m") or ""
                minute = 30 if m_raw == "半" else (int(m_raw) if m_raw.isdigit() else 0)
                if h is not None and 0 <= h <= 24 and 0 <= minute < 60:
                    tokens.append({
                        "start": m.start(),
                        "end": m.end(),
                        "h": h,
                        "m": minute,
                        "kind": "explicit",
                        "pm": tok_pm,
                        "am": tok_am,
                    })
            elif m.group("c_h") is not None:
                h = int(m.group("c_h"))
                minute = int(m.group("c_m"))
                sep = m.group("c_sep")
                # 点号形式（如 9.30、14.00）只认整点与半点，避免把日期 11.18 当成时间
                if sep == "." and minute not in (0, 30):
                    continue
                if 0 <= h <= 24 and 0 <= minute < 60:
                    tokens.append({
                        "start": m.start(),
                        "end": m.end(),
                        "h": h,
                        "m": minute,
                        "kind": "explicit",
                        "pm": tok_pm,
                        "am": tok_am,
                    })
            elif m.group("m_tok") is not None:
                raw = m.group("m_tok")
                if len(raw) == 4:
                    h, minute = int(raw[:2]), int(raw[2:])
                    valid = (6 <= h <= 23) and (minute in (0, 30))
                else:
                    h, minute = int(raw[:1]), int(raw[1:])
                    valid = (1 <= h <= 9) and (minute in (0, 30))
                if valid:
                    tokens.append({
                        "start": m.start(),
                        "end": m.end(),
                        "h": h,
                        "m": minute,
                        "kind": "compact4" if len(raw) == 4 else "compact3",
                        "pm": tok_pm,
                        "am": tok_am,
                    })
            elif m.group("b_h") is not None:
                h = int(m.group("b_h"))
                if 1 <= h <= 24:
                    tokens.append({
                        "start": m.start(),
                        "end": m.end(),
                        "h": h,
                        "m": 0,
                        "kind": "bare",
                        "pm": tok_pm,
                        "am": tok_am,
                    })

        return tokens

    def _extract_all_time_ranges(self, msg: str, post: str = "") -> List[Tuple[int, int, str]]:
        """
        从单个片段中提取所有时间段：
        基于统一 Time-Token 配对，一次打通 9-1030、8-930、2-330、1230-2、1530-18、9.30-12、
        15:30-18:00、14:00到下午6点、两点到六点、1400-1530 1530 1800 等全部混合格式。
        """
        # D5图书室弹性时段
        if post == "D5图书室" or "D5" in msg.upper():
            if any(k in msg for k in ["上午", "早上", "早班"]):
                return [(-1, -1, "上午")]
            if any(k in msg for k in ["晚上", "晚班", "今晚", "明晚"]):
                return [(-2, -2, "晚上")]

        is_pm_context = any(k in msg for k in ["下午", "晚上", "今晚", "明晚", "后晚", "晚班", "夜班", "晚"])
        is_am_context = any(k in msg for k in ["早上", "早班", "上午", "今早", "明早", "早"]) and not is_pm_context

        valid_short_pairs = {
            (8, 9), (9, 12), (12, 14), (14, 16), (16, 18), (14, 18),
            (18, 20), (20, 22), (20, 23), (18, 22), (18, 23),
            (12, 2), (12, 4), (2, 4), (2, 6), (4, 6),
            (6, 8), (6, 10), (6, 11), (8, 10), (8, 11),
            (9, 11), (10, 12),
        }

        tokens = self._parse_single_time_token(msg)
        results: List[Tuple[int, int, str]] = []
        conn_re = re.compile(r"^\s*[-~～—–−至到]+\s*$")
        space_re = re.compile(r"^\s+$")

        i = 0
        while i + 1 < len(tokens):
            t1 = tokens[i]
            t2 = tokens[i + 1]
            gap = msg[t1["end"]:t2["start"]]

            is_dash_conn = bool(conn_re.match(gap))
            both_compact = (t1["kind"] == "compact4" and t2["kind"] == "compact4")
            is_space_conn = bool(both_compact and space_re.match(gap))

            if not (is_dash_conn or is_space_conn):
                i += 1
                continue

            # 若两端都是纯1~2位无修饰短数字（如 8-10, 2-4），必须命中合法短数字集合，避免把普通数字区间当成时间
            if t1["kind"] == "bare" and t2["kind"] == "bare" and not (t1["pm"] or t1["am"] or t2["pm"] or t2["am"]):
                if (t1["h"], t2["h"]) not in valid_short_pairs:
                    i += 1
                    continue

            t_start, t_end, t_label = self._normalize_library_hours(
                t1["h"],
                t1["m"],
                t2["h"],
                t2["m"],
                is_pm_context=is_pm_context,
                is_am_context=is_am_context,
                t1_pm=t1["pm"],
                t1_am=t1["am"],
                t2_pm=t2["pm"],
                t2_am=t2["am"],
            )
            # 过滤明显不合理的长区间（如日期差或误匹配），图书馆单条连班跨度不超过 10 小时
            if 30 <= (t_end - t_start) <= 600 and 420 <= t_start and t_end <= 1440:
                results.append((t_start, t_end, t_label))
                i += 2
            else:
                i += 1

        return results

    def _normalize_library_hours(
        self,
        h1: int,
        m1: int,
        h2: int,
        m2: int,
        is_pm_context: bool,
        is_am_context: bool = False,
        t1_pm: bool = False,
        t1_am: bool = False,
        t2_pm: bool = False,
        t2_am: bool = False,
    ) -> Tuple[int, int, str]:
        """
        根据图书馆真实开放时段（08:00 ~ 23:00）自动消除 12/24 小时制及混合简写歧义：
        - 1~7 点必然为下午/傍晚 13~19 点（如 2-330 -> 14:00-15:30, 6-8 -> 18:00-20:00）
        - 8-10 / 8-11 必然为晚班 20-22 / 20-23（早班只有 8:00-9:00 或 8:00-9:30）
        - 1230-2 / 12:30-2:00 结束小时在 1~7 自动转 14:00
        - 只有带“晚/下午”局部修饰或整句晚间上下文且符合晚班习惯时才对 8/9/10 点 +12
        """
        if 1 <= h1 <= 7:
            h1 += 12
        elif h1 == 8:
            is_evening_pair = (m1 == 0 and m2 == 0 and h2 in (10, 11) and not is_am_context and not t1_am)
            if t1_pm or (is_pm_context and not t1_am and not is_am_context) or is_evening_pair:
                h1 += 12
        elif 9 <= h1 <= 11:
            if t1_pm:
                h1 += 12

        if 1 <= h2 <= 7:
            h2 += 12
        elif t2_pm and h2 <= 11:
            h2 += 12
        elif h1 >= 12 and h2 < 12:
            h2 += 12

        return h1 * 60 + m1, h2 * 60 + m2, f"{h1:02d}:{m1:02d}-{h2:02d}:{m2:02d}"
