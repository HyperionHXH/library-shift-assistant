# -*- coding: utf-8 -*-
"""
PC端微信群/聊天窗口实时监控与极速自动回复模块 (wechat_monitor.py)
双引擎自动适配微信 3.x (WeChat.exe) 与微信 4.x (Weixin.exe / Qt51514QWindowIcon)：
1. 优先尝试 Windows UIAutomation 原生树读取；
2. 当检测到微信 4.x (Qt51514QWindowIcon) 默认隐藏 UIA 控件树时，自动切换为：
   【DWM 后台窗口渲染 (PrintWindow) + 自适应气泡分割 + Windows 原生中文 OCR 引擎】！
   （严格声明 64 位 Win32 句柄类型，彻底修复 64 位 GDI 句柄溢出并支持微信被遮挡时的后台读取）
3. 自动识别聊天区域分割线（同时兼容“微信主窗口”与“独立弹出聊天窗口”）。
4. 【防自己接自己】绿色右侧气泡 / 右对齐消息直接标为 is_self=True，绝不误接本人测试消息！
5. 采用滑动尾缀窗口 + OCR 归一化防抖比对，启动时不误回旧消息，也不因 OCR 微小抖动重复触发。
6. 精确定位底部输入框并极速激活窗口粘贴发送回复。
"""

from __future__ import annotations
import asyncio
import ctypes
from ctypes import wintypes
import datetime
import difflib
import os
import re
import time
import threading
from collections import deque
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Tuple

from PIL import Image, ImageEnhance, ImageGrab, ImageOps

# ==============================================================================
# 64位安全的 Win32 API 声明（全部显式指定 argtypes 与 restype，杜绝 OverflowError）
# ==============================================================================
user32 = ctypes.WinDLL("user32", use_last_error=True)
gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
PW_RENDERFULLCONTENT = 0x00000002
DIB_RGB_COLORS = 0
SW_RESTORE = 9

user32.IsWindow.argtypes = [wintypes.HWND]
user32.IsWindow.restype = wintypes.BOOL
user32.IsWindowVisible.argtypes = [wintypes.HWND]
user32.IsWindowVisible.restype = wintypes.BOOL
user32.IsIconic.argtypes = [wintypes.HWND]
user32.IsIconic.restype = wintypes.BOOL
user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
user32.GetWindowRect.restype = wintypes.BOOL
user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
user32.GetWindowTextLengthW.restype = ctypes.c_int
user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetWindowTextW.restype = ctypes.c_int
user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetClassNameW.restype = ctypes.c_int
user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user32.GetWindowThreadProcessId.restype = wintypes.DWORD
user32.GetWindowDC.argtypes = [wintypes.HWND]
user32.GetWindowDC.restype = wintypes.HDC
user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
user32.ReleaseDC.restype = ctypes.c_int
user32.PrintWindow.argtypes = [wintypes.HWND, wintypes.HDC, wintypes.UINT]
user32.PrintWindow.restype = wintypes.BOOL
user32.SetForegroundWindow.argtypes = [wintypes.HWND]
user32.SetForegroundWindow.restype = wintypes.BOOL
user32.BringWindowToTop.argtypes = [wintypes.HWND]
user32.BringWindowToTop.restype = wintypes.BOOL
user32.GetForegroundWindow.argtypes = []
user32.GetForegroundWindow.restype = wintypes.HWND
user32.AttachThreadInput.argtypes = [wintypes.DWORD, wintypes.DWORD, wintypes.BOOL]
user32.AttachThreadInput.restype = wintypes.BOOL
user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
user32.ShowWindow.restype = wintypes.BOOL
user32.OpenClipboard.argtypes = [wintypes.HWND]
user32.OpenClipboard.restype = wintypes.BOOL
user32.EmptyClipboard.argtypes = []
user32.EmptyClipboard.restype = wintypes.BOOL
user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
user32.SetClipboardData.restype = wintypes.HANDLE
user32.CloseClipboard.argtypes = []
user32.CloseClipboard.restype = wintypes.BOOL
user32.GetSystemMetrics.argtypes = [ctypes.c_int]
user32.GetSystemMetrics.restype = ctypes.c_int
user32.EnumWindows.argtypes = [ctypes.c_void_p, wintypes.LPARAM]
user32.EnumWindows.restype = wintypes.BOOL

gdi32.CreateCompatibleDC.argtypes = [wintypes.HDC]
gdi32.CreateCompatibleDC.restype = wintypes.HDC
gdi32.CreateCompatibleBitmap.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int]
gdi32.CreateCompatibleBitmap.restype = wintypes.HBITMAP
gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
gdi32.SelectObject.restype = wintypes.HGDIOBJ
gdi32.GetDIBits.argtypes = [wintypes.HDC, wintypes.HBITMAP, wintypes.UINT, wintypes.UINT, ctypes.c_void_p, ctypes.c_void_p, wintypes.UINT]
gdi32.GetDIBits.restype = ctypes.c_int
gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
gdi32.DeleteObject.restype = wintypes.BOOL
gdi32.DeleteDC.argtypes = [wintypes.HDC]
gdi32.DeleteDC.restype = wintypes.BOOL

kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
kernel32.OpenProcess.restype = wintypes.HANDLE
kernel32.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
kernel32.CloseHandle.restype = wintypes.BOOL
kernel32.GetCurrentThreadId.argtypes = []
kernel32.GetCurrentThreadId.restype = wintypes.DWORD
kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalLock.restype = ctypes.c_void_p
kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalUnlock.restype = wintypes.BOOL


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", wintypes.DWORD),
        ("biWidth", wintypes.LONG),
        ("biHeight", wintypes.LONG),
        ("biPlanes", wintypes.WORD),
        ("biBitCount", wintypes.WORD),
        ("biCompression", wintypes.DWORD),
        ("biSizeImage", wintypes.DWORD),
        ("biXPelsPerMeter", wintypes.LONG),
        ("biYPelsPerMeter", wintypes.LONG),
        ("biClrUsed", wintypes.DWORD),
        ("biClrImportant", wintypes.DWORD),
    ]


class BITMAPINFO(ctypes.Structure):
    _fields_ = [
        ("bmiHeader", BITMAPINFOHEADER),
        ("bmiColors", wintypes.DWORD * 3),
    ]


def _get_process_name_by_hwnd(hwnd: int) -> str:
    """获取窗口所属进程的可执行文件名（如 WeChat.exe / Weixin.exe）"""
    try:
        pid = wintypes.DWORD(0)
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if not pid.value:
            return ""
        h_proc = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid.value)
        if not h_proc:
            return ""
        try:
            buf = ctypes.create_unicode_buffer(512)
            size = wintypes.DWORD(512)
            if kernel32.QueryFullProcessImageNameW(h_proc, 0, buf, ctypes.byref(size)):
                return os.path.basename(buf.value).lower()
        finally:
            kernel32.CloseHandle(h_proc)
    except Exception:
        pass
    return ""


def set_clipboard_text(text: str) -> bool:
    """使用 Win32 API 极速写入 Unicode 文本到剪贴板"""
    try:
        if not user32.OpenClipboard(0):
            return False
        user32.EmptyClipboard()
        data = text.encode("utf-16-le") + b"\x00\x00"
        h_mem = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data))
        if not h_mem:
            user32.CloseClipboard()
            return False
        ptr = kernel32.GlobalLock(h_mem)
        if not ptr:
            user32.CloseClipboard()
            return False
        ctypes.memmove(ptr, data, len(data))
        kernel32.GlobalUnlock(h_mem)
        user32.SetClipboardData(CF_UNICODETEXT, h_mem)
        user32.CloseClipboard()
        return True
    except Exception:
        try:
            user32.CloseClipboard()
        except Exception:
            pass
        return False


def _activate_window(hwnd: int) -> None:
    """可靠将目标窗口激活并置于前台（即使之前被其他窗口遮挡）"""
    if not hwnd or not user32.IsWindow(hwnd):
        return
    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, SW_RESTORE)

    fg_hwnd = user32.GetForegroundWindow()
    if fg_hwnd == hwnd:
        return

    cur_tid = kernel32.GetCurrentThreadId()
    fg_tid = user32.GetWindowThreadProcessId(fg_hwnd, None) if fg_hwnd else 0
    attached = False
    try:
        if fg_tid and fg_tid != cur_tid:
            attached = bool(user32.AttachThreadInput(cur_tid, fg_tid, True))
        user32.BringWindowToTop(hwnd)
        user32.SetForegroundWindow(hwnd)
    finally:
        if attached and fg_tid:
            user32.AttachThreadInput(cur_tid, fg_tid, False)


def capture_window_image(hwnd: int) -> Tuple[Optional[Image.Image], Tuple[int, int, int, int], str]:
    """
    截取指定 HWND 窗口的图像，返回 (PIL.Image, (win_left, win_top, win_right, win_bottom), capture_method)。
    先尝试 PrintWindow(PW_RENDERFULLCONTENT)（即使微信被其他窗口完全遮挡也能直接读取 DWM 后台渲染画面）；
    若全黑则回退到屏幕可见区域抓取。
    """
    if not hwnd or not user32.IsWindow(hwnd) or user32.IsIconic(hwnd):
        return None, (0, 0, 0, 0), "窗口不可用或已最小化"

    rect = wintypes.RECT()
    if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        return None, (0, 0, 0, 0), "获取窗口坐标失败"
    w = rect.right - rect.left
    h = rect.bottom - rect.top
    if w < 100 or h < 100:
        return None, (rect.left, rect.top, rect.right, rect.bottom), "窗口尺寸过小"

    img_pw: Optional[Image.Image] = None
    hdc_win = user32.GetWindowDC(hwnd)
    if hdc_win:
        hdc_mem = None
        h_bmp = None
        old_obj = None
        try:
            hdc_mem = gdi32.CreateCompatibleDC(hdc_win)
            if hdc_mem:
                h_bmp = gdi32.CreateCompatibleBitmap(hdc_win, w, h)
                if h_bmp:
                    old_obj = gdi32.SelectObject(hdc_mem, h_bmp)
                    for pw_flag in (PW_RENDERFULLCONTENT, 0):
                        ok = user32.PrintWindow(hwnd, hdc_mem, pw_flag)
                        if not ok:
                            continue
                        bmi = BITMAPINFO()
                        bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
                        bmi.bmiHeader.biWidth = w
                        bmi.bmiHeader.biHeight = -h
                        bmi.bmiHeader.biPlanes = 1
                        bmi.bmiHeader.biBitCount = 32
                        bmi.bmiHeader.biCompression = 0
                        buf = ctypes.create_string_buffer(w * h * 4)
                        bits = gdi32.GetDIBits(hdc_mem, h_bmp, 0, h, buf, ctypes.byref(bmi), DIB_RGB_COLORS)
                        if bits:
                            candidate = Image.frombuffer("RGBA", (w, h), buf, "raw", "BGRA", 0, 1).convert("RGB")
                            extrema = candidate.getextrema()
                            # 只要整张图非纯黑且存在实际像素变化，即为有效渲染图像
                            if any(hi > 20 for _lo, hi in extrema) and any(hi - lo > 10 for lo, hi in extrema):
                                img_pw = candidate
                                break
        except Exception:
            pass
        finally:
            if hdc_mem and old_obj:
                gdi32.SelectObject(hdc_mem, old_obj)
            if h_bmp:
                gdi32.DeleteObject(h_bmp)
            if hdc_mem:
                gdi32.DeleteDC(hdc_mem)
            user32.ReleaseDC(hwnd, hdc_win)

    if img_pw is not None:
        return img_pw, (rect.left, rect.top, rect.right, rect.bottom), "PrintWindow后台渲染(支持遮挡)"

    # 回退到可见区域屏幕抓取
    try:
        scr_w = user32.GetSystemMetrics(0)
        scr_h = user32.GetSystemMetrics(1)
        x0 = max(0, rect.left)
        y0 = max(0, rect.top)
        x1 = min(scr_w, rect.right)
        y1 = min(scr_h, rect.bottom)
        if x1 - x0 >= 80 and y1 - y0 >= 80:
            grabbed = ImageGrab.grab(bbox=(x0, y0, x1, y1), all_screens=True).convert("RGB")
            return grabbed, (x0, y0, x1, y1), "屏幕可见区抓取(需避免完全遮挡)"
    except Exception:
        pass

    return None, (rect.left, rect.top, rect.right, rect.bottom), "画面获取失败"


# ==============================================================================
# Windows 11 / 10 原生 WinRT OCR 与微信 4.0 气泡视觉识别
# ==============================================================================
class WinRTOcrHelper:
    """线程安全的 Windows 原生中文 OCR 封装"""

    def __init__(self):
        self._engine = None
        self._init_attempted = False
        self._loop = None

    def _ensure_engine(self) -> bool:
        if self._engine is not None:
            return True
        if self._init_attempted:
            return False
        self._init_attempted = True
        try:
            from winrt.windows.media.ocr import OcrEngine
            from winrt.windows.globalization import Language

            eng = OcrEngine.try_create_from_language(Language("zh-Hans-CN"))
            if eng is None:
                eng = OcrEngine.try_create_from_user_profile_languages()
            self._engine = eng
            return self._engine is not None
        except Exception:
            return False

    def is_available(self) -> bool:
        return self._ensure_engine()

    async def _recognize_async(
        self,
        img: Image.Image,
        scale: float = 3.0,
        pad_color: Optional[Tuple[int, int, int]] = None,
    ) -> List[Tuple[float, float, float, float, str]]:
        from winrt.windows.graphics.imaging import SoftwareBitmap, BitmapPixelFormat, BitmapAlphaMode
        from winrt.windows.storage.streams import DataWriter

        work = img.convert("RGB")
        # 边缘留白 16px，避免紧密裁剪时边缘汉字笔画被截断失真
        if pad_color is not None:
            work = ImageOps.expand(work, border=16, fill=pad_color)
        if scale != 1.0:
            work = work.resize((max(1, int(work.width * scale)), max(1, int(work.height * scale))), Image.Resampling.LANCZOS)
        work = ImageEnhance.Contrast(work).enhance(1.35)
        rgba = work.convert("RGBA")

        writer = DataWriter()
        writer.write_bytes(rgba.tobytes())
        sb = SoftwareBitmap.create_copy_from_buffer(
            writer.detach_buffer(),
            BitmapPixelFormat.RGBA8,
            rgba.width,
            rgba.height,
            BitmapAlphaMode.PREMULTIPLIED,
        )
        res = await self._engine.recognize_async(sb)
        rows = []
        for line in res.lines:
            bbs = [w.bounding_rect for w in line.words]
            if bbs:
                x0 = min(r.x for r in bbs) / scale
                y0 = min(r.y for r in bbs) / scale
                x1 = max(r.x + r.width for r in bbs) / scale
                y1 = max(r.y + r.height for r in bbs) / scale
            else:
                x0 = y0 = x1 = y1 = 0.0
            rows.append((y0, x0, x1, y1, line.text))
        return rows

    def recognize_text_lines(
        self,
        img: Image.Image,
        scale: float = 3.0,
        pad_color: Optional[Tuple[int, int, int]] = None,
    ) -> List[Tuple[float, float, float, float, str]]:
        if not self._ensure_engine():
            return []
        try:
            if self._loop is None or self._loop.is_closed():
                self._loop = asyncio.new_event_loop()
            rows = self._loop.run_until_complete(self._recognize_async(img, scale=scale, pad_color=pad_color))
        except Exception:
            return []

        rows.sort(key=lambda r: (r[0], r[1]))
        grouped: List[List] = []
        for y0, x0, x1, y1, txt in rows:
            if grouped and abs(y0 - grouped[-1][0]) < 8.0:
                grouped[-1][1].append((x0, txt))
                grouped[-1][2] = min(grouped[-1][2], x0)
                grouped[-1][3] = max(grouped[-1][3], x1)
                grouped[-1][4] = max(grouped[-1][4], y1)
            else:
                grouped.append([y0, [(x0, txt)], x0, x1, y1])

        out: List[Tuple[float, float, float, float, str]] = []
        for y0, chunks, x0, x1, y1 in grouped:
            chunks.sort(key=lambda c: c[0])
            joined = " ".join(c[1] for c in chunks)
            cleaned = clean_ocr_text(joined)
            if cleaned:
                out.append((y0, x0, x1, y1, cleaned))
        return out


def clean_ocr_text(s: str) -> str:
    """清理 WinRT OCR 对中英文数字混排产生的多余空格及连字符误识别"""
    s = re.sub(r"(?<=[\u4e00-\u9fff:：\-~～，,、;；（\(）\)@])\s+(?=[\u4e00-\u9fff0-9:：\-~～，,、;；（\(）\)])", "", s)
    s = re.sub(r"(?<=\d)\s+(?=[:：\.\-~～\u4e00-\u9fff])", "", s)
    s = re.sub(r"(?<!\d)(\d)\s+(\d)(?=[:：\-\u4e00-\u9fff]|$)", r"\1\2", s)
    s = re.sub(r"(?<=\d)\s*[·•．]\s*(?=\d)", "-", s)
    s = re.sub(r"(?<=[:：\d]00)\s*[一—]\s*(?=\d)", "-", s)
    s = re.sub(r"(?<=[:：\d]30)\s*[一—]\s*(?=\d)", "-", s)
    return s.strip()


def _join_bubble_ocr_lines(
    lines: List[Tuple[float, float, float, float, str]],
    bubble_width: int,
) -> str:
    """
    智能合并气泡内的多行 OCR 文本：
    - 若上一行到达气泡右侧折行边界，或跨行把同一个词拆断（如“周\\n四”“2-\\n330”“三楼\\n书库”），
      直接无缝拼接避免误拆成两段；
    - 若是用户主动按回车换行分条发送的短行，保留换行分隔。
    """
    if not lines:
        return ""
    if len(lines) == 1:
        return lines[0][4].strip()

    max_x1 = max(ln[2] for ln in lines)
    parts: List[str] = [lines[0][4].strip()]
    for i in range(1, len(lines)):
        prev_txt = parts[-1]
        prev_x1 = lines[i - 1][2]
        cur_txt = lines[i][4].strip()
        if not cur_txt:
            continue

        # 规则1：跨行把词组、日期、时间、岗位拆断时强制无缝拼接
        broken_token = bool(
            (re.search(r"(?:周|星期|礼拜)$", prev_txt) and re.match(r"^[一二三四五六日天末1-7]", cur_txt))
            or (re.search(r"[上下早中晚今明后大]$", prev_txt) and re.match(r"^[午上早晚天]", cur_txt))
            or (re.search(r"[一二两三四五六七八九十\d]$", prev_txt) and re.match(r"^[楼层]", cur_txt))
            or (prev_txt.endswith(("服", "服务", "书", "图", "图书")) and re.match(r"^[务台库书室]", cur_txt))
            or (re.search(r"[求代帮顶出可]$", prev_txt) and re.match(r"^[替班代]", cur_txt))
            or (re.search(r"[\d\-~～—–−至到:：\.]$", prev_txt) and re.match(r"^[\d\-~～—–−至到:：]", cur_txt))
        )
        # 规则2：气泡右边缘自动折行（非句末换行）也合并，具体分段交给 _split_into_shift_segments 识别
        soft_wrapped = (
            max_x1 >= max(60, bubble_width - 28)
            and prev_x1 >= max_x1 - 18
            and not prev_txt.endswith(("。", "！", "!", "？", "?"))
        )
        if broken_token or soft_wrapped:
            parts[-1] = prev_txt + cur_txt
        else:
            parts.append(cur_txt)

    return "\n".join(parts).strip()


def _locate_chat_pane_rect(win_img: Image.Image, title: str = "") -> Tuple[int, int, int, int, int]:
    """自动定位微信窗口中的【右侧聊天消息区】及【底部输入框 Y 坐标】"""
    w, h = win_img.size
    rgb = win_img.convert("RGB")
    px = rgb.tobytes()
    stride = w * 3

    chat_x0 = 4
    if w >= 620:
        y_samples = range(int(h * 0.15), int(h * 0.72), max(1, h // 40))
        best_x = None
        best_score = 0
        for x in range(int(w * 0.22), int(w * 0.54)):
            border_cnt = 0
            bright_right = 0
            for y in y_samples:
                idx = y * stride + x * 3
                r, g, b = px[idx], px[idx + 1], px[idx + 2]
                if 205 <= r <= 242 and abs(r - g) <= 4 and abs(g - b) <= 4:
                    border_cnt += 1
                idx_r = y * stride + min(w - 1, x + 8) * 3
                rr, gg, bb = px[idx_r], px[idx_r + 1], px[idx_r + 2]
                if (rr + gg + bb) // 3 >= 242:
                    bright_right += 1
            if border_cnt >= len(y_samples) * 0.75 and bright_right >= len(y_samples) * 0.65:
                score = border_cnt + bright_right
                if score > best_score:
                    best_score = score
                    best_x = x
        if best_x is not None:
            chat_x0 = best_x + 2

    chat_x1 = w - 8
    x_samples = range(chat_x0 + 20, max(chat_x0 + 21, chat_x1 - 20), max(1, (chat_x1 - chat_x0) // 30))

    chat_y0 = min(64, int(h * 0.12))
    for y in range(36, min(120, int(h * 0.22))):
        cnt = 0
        for x in x_samples:
            idx = y * stride + x * 3
            r, g, b = px[idx], px[idx + 1], px[idx + 2]
            if 205 <= r <= 244 and abs(r - g) <= 4 and abs(g - b) <= 4:
                cnt += 1
        if cnt >= len(x_samples) * 0.80:
            chat_y0 = y + 2
            break

    chat_y1 = int(h * 0.76)
    for y in range(int(h * 0.58), min(h - 4, int(h * 0.88))):
        cnt = 0
        bright_neighbors = 0
        for x in x_samples:
            idx = y * stride + x * 3
            r, g, b = px[idx], px[idx + 1], px[idx + 2]
            if 205 <= r <= 244 and abs(r - g) <= 4 and abs(g - b) <= 4:
                cnt += 1
                idx_up = max(0, y - 3) * stride + x * 3
                idx_dn = min(h - 1, y + 3) * stride + x * 3
                if px[idx_up] >= 245 and px[idx_dn] >= 245:
                    bright_neighbors += 1
        if cnt >= len(x_samples) * 0.80 and bright_neighbors >= len(x_samples) * 0.60:
            chat_y1 = y - 2
            break

    if chat_y1 <= chat_y0 + 80:
        chat_y0 = int(h * 0.10)
        chat_y1 = int(h * 0.76)

    input_area_h = max(40, h - chat_y1)
    input_click_y = min(h - 18, chat_y1 + max(42, int(input_area_h * 0.55)))
    return chat_x0, chat_y0, chat_x1, chat_y1, input_click_y


def detect_chat_bubbles_adaptive(chat_img: Image.Image) -> List[Tuple[int, int, int, int, bool]]:
    """自适应检测微信 4.0 (灰白气泡 #EEEEF0 / #FAFAFA 背景) 与 微信 3.x (白气泡 / #EDEDED 背景) 及绿色己方气泡。"""
    w, h = chat_img.size
    if w < 60 or h < 60:
        return []
    sw, sh = max(1, w // 2), max(1, h // 2)
    small = chat_img.resize((sw, sh), Image.Resampling.NEAREST).convert("RGB")
    px = small.tobytes()
    stride = sw * 3

    bg_lums = []
    for y in range(int(sh * 0.08), int(sh * 0.92), max(1, sh // 25)):
        for x in (int(sw * 0.04), int(sw * 0.50), int(sw * 0.86)):
            idx = y * stride + x * 3
            r, g, b = px[idx], px[idx + 1], px[idx + 2]
            if abs(r - g) <= 6 and abs(g - b) <= 6:
                bg_lums.append((r + g + b) // 3)
    bg_lum = max(bg_lums) if bg_lums else 250
    is_wx4_light = bg_lum >= 247
    is_dark_mode = bg_lum < 80

    row_info = []
    x_start = int(sw * 0.08)
    x_end = int(sw * 0.96)
    for y in range(sh):
        row_off = y * stride
        inc_cnt = 0
        inc_min, inc_max = sw, -1
        slf_cnt = 0
        slf_min, slf_max = sw, -1
        for x in range(x_start, x_end):
            idx = row_off + x * 3
            r, g, b = px[idx], px[idx + 1], px[idx + 2]
            is_green = (g >= 110 and (g - r) >= 24 and (g - b) >= 20)
            if is_green:
                slf_cnt += 1
                if x < slf_min:
                    slf_min = x
                if x > slf_max:
                    slf_max = x
                continue

            if is_dark_mode:
                lum = (r + g + b) // 3
                is_in = (bg_lum + 10 <= lum <= bg_lum + 50) and abs(r - g) <= 8
            elif is_wx4_light:
                is_in = (224 <= r <= 243 and 224 <= g <= 243 and 226 <= b <= 246 and abs(r - g) <= 6)
            else:
                is_in = (r >= 251 and g >= 251 and b >= 251)

            if is_in:
                inc_cnt += 1
                if x < inc_min:
                    inc_min = x
                if x > inc_max:
                    inc_max = x
        row_info.append((inc_min, inc_max, inc_cnt, slf_min, slf_max, slf_cnt))

    bubbles: List[Tuple[int, int, int, int, bool]] = []

    def extract_bands(is_green: bool):
        in_band = False
        b_y0 = 0
        b_x0, b_x1 = sw, 0
        gap = 0
        for y, info in enumerate(row_info):
            xmin, xmax, cnt = (info[3], info[4], info[5]) if is_green else (info[0], info[1], info[2])
            if cnt >= 10:
                if not in_band:
                    in_band = True
                    b_y0 = y
                    b_x0, b_x1 = xmin, xmax
                else:
                    if xmin < b_x0:
                        b_x0 = xmin
                    if xmax > b_x1:
                        b_x1 = xmax
                gap = 0
            else:
                if in_band:
                    gap += 1
                    if gap >= 3:
                        b_y1 = y - gap
                        if (b_y1 - b_y0) >= 8:
                            bubbles.append((b_y0 * 2, (b_y1 + 1) * 2, max(0, b_x0 * 2 - 4), min(w, (b_x1 + 1) * 2 + 4), is_green))
                        in_band = False
                        gap = 0
        if in_band:
            b_y1 = sh - 1 - gap
            if (b_y1 - b_y0) >= 8:
                bubbles.append((b_y0 * 2, (b_y1 + 1) * 2, max(0, b_x0 * 2 - 4), min(w, (b_x1 + 1) * 2 + 4), is_green))

    extract_bands(False)
    extract_bands(True)
    bubbles.sort(key=lambda b: b[0])
    return bubbles


# ==============================================================================
# 窗口扫描与消息比对数据结构
# ==============================================================================
@dataclass
class WeChatWindowInfo:
    hwnd: int
    title: str
    class_name: str
    process_name: str = ""

    def display_label(self) -> str:
        is_single = (
            self.class_name in ("ChatWnd", "mmui::ChatSingleWindow")
            or (self.title and self.title not in ("微信", "WeChat", "Weixin"))
        )
        kind = "独立聊天窗" if is_single else "微信主窗口"
        return f"{self.title or '微信'} [{kind} | HWND:{self.hwnd}]"


@dataclass
class CapturedMessage:
    fingerprint: str
    sender: str
    content: str
    is_self: bool = False


def list_wechat_windows() -> List[WeChatWindowInfo]:
    """扫描当前桌面上所有可见的微信窗口（支持微信 3.x 与微信 4.x）"""
    results: List[WeChatWindowInfo] = []
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)

    def enum_cb(hwnd, _lparam):
        if not user32.IsWindowVisible(hwnd):
            return True

        rect = wintypes.RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        if (rect.right - rect.left) < 200 or (rect.bottom - rect.top) < 200:
            return True

        length = user32.GetWindowTextLengthW(hwnd)
        title_buf = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, title_buf, length + 1)
        title = title_buf.value.strip()

        cls_buf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, cls_buf, 256)
        cls_name = cls_buf.value.strip()

        proc_name = _get_process_name_by_hwnd(int(hwnd))
        is_wx_proc = proc_name in ("wechat.exe", "weixin.exe")
        is_wx_cls = (
            cls_name in ("WeChatMainWndForPC", "ChatWnd", "mmui::MainWindow", "mmui::ChatSingleWindow", "Weixin")
            or "WeChat" in cls_name
            or ("Qt" in cls_name and title == "微信")
        )

        if is_wx_proc or is_wx_cls:
            if cls_name in ("ImagePreviewWnd", "CefWebViewWnd", "mmui::ImagePreviewWindow"):
                return True
            results.append(
                WeChatWindowInfo(
                    hwnd=int(hwnd),
                    title=title or "微信",
                    class_name=cls_name,
                    process_name=proc_name,
                )
            )
        return True

    cb_ref = WNDENUMPROC(enum_cb)
    user32.EnumWindows(cb_ref, 0)
    results.sort(key=lambda w: (0 if w.class_name in ("ChatWnd", "mmui::ChatSingleWindow") else 1, w.hwnd))
    return results


def _iter_controls_bfs(root_ctrl, max_depth: int = 10, max_nodes: int = 400):
    queue = deque([(root_ctrl, 0)])
    visited = 0
    while queue and visited < max_nodes:
        ctrl, depth = queue.popleft()
        visited += 1
        yield ctrl, depth
        if depth < max_depth:
            try:
                for ch in ctrl.GetChildren():
                    queue.append((ch, depth + 1))
            except Exception:
                pass


def _norm_msg_for_match(s: str) -> str:
    s = s.replace("：", ":").replace("，", ",").replace("（", "(").replace("）", ")")
    return re.sub(r"[\s·•\.\-_]+", "", s)


def _msgs_similar(a: CapturedMessage, b: CapturedMessage) -> bool:
    if a.is_self != b.is_self:
        return False
    if a.fingerprint == b.fingerprint:
        return True
    na, nb = _norm_msg_for_match(a.content), _norm_msg_for_match(b.content)
    if not na or not nb:
        return False
    if na == nb:
        return True
    if min(len(na), len(nb)) >= 6:
        return difflib.SequenceMatcher(None, na, nb).ratio() >= 0.88
    return False


def _compute_new_tail_items(
    prev_tail: List[CapturedMessage],
    cur_tail: List[CapturedMessage],
) -> List[CapturedMessage]:
    if not cur_tail or not prev_tail:
        return []

    max_k = min(len(prev_tail), len(cur_tail))
    for k in range(max_k, 0, -1):
        ok = True
        for i in range(k):
            if not _msgs_similar(prev_tail[len(prev_tail) - k + i], cur_tail[i]):
                ok = False
                break
        if ok:
            return cur_tail[k:]

    if any(_msgs_similar(cur_tail[-1], old) for old in prev_tail[-3:]):
        return []
    return cur_tail[-1:]


class WeChatMonitor:
    """PC微信群/聊天窗后台监控线程控制器（UIA + PrintWindow 后台 OCR 双引擎）"""

    def __init__(
        self,
        on_new_message: Callable[[str, str, bool], None],
        on_log: Callable[[str], None],
        on_status_change: Optional[Callable[[bool, str], None]] = None,
    ):
        self.on_new_message = on_new_message
        self.on_log = on_log
        self.on_status_change = on_status_change

        self.is_running: bool = False
        self._stop_event = threading.Event()
        self._thread: Optional[threading.Thread] = None

        self.target_hwnd: Optional[int] = None
        self.target_group_keyword: str = ""
        self.poll_interval_sec: float = 0.40
        self.auto_send_enabled: bool = True
        self.cooldown_sec: float = 2.5
        self._last_send_time: float = 0.0

        self._ocr = WinRTOcrHelper()
        self._last_engine_used: str = ""
        self._cached_input_abs_xy: Optional[Tuple[int, int]] = None

        self._recent_sent_replies: List[Tuple[str, float]] = []
        self._recent_processed: Dict[str, float] = {}
        self._last_tail: List[CapturedMessage] = []

    def start(
        self,
        target_hwnd: Optional[int] = None,
        target_group_keyword: str = "",
        poll_interval_sec: float = 0.40,
        auto_send_enabled: bool = True,
    ) -> Tuple[bool, str]:
        if self.is_running:
            return False, "监控程序已在运行中"

        self.target_hwnd = target_hwnd
        self.target_group_keyword = target_group_keyword.strip()
        self.poll_interval_sec = max(0.15, float(poll_interval_sec))
        self.auto_send_enabled = auto_send_enabled
        self._last_tail = []
        self._recent_sent_replies.clear()
        self._recent_processed.clear()
        self._last_engine_used = ""
        self._stop_event.clear()

        self._thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.is_running = True
        self._thread.start()
        if self.on_status_change:
            self.on_status_change(True, "正在监控微信消息...")
        return True, "监控已启动"

    def stop(self) -> None:
        self._stop_event.set()
        self.is_running = False
        if self.on_status_change:
            self.on_status_change(False, "监控已停止")

    def inspect_wechat_window(self, target_hwnd: Optional[int] = None, keyword: str = "") -> str:
        try:
            import uiautomation as auto
            with auto.UIAutomationInitializerInThread():
                self.target_hwnd = target_hwnd
                self.target_group_keyword = keyword.strip()
                wx_ctrl = self._find_target_control(auto)
                if not wx_ctrl or not self.target_hwnd:
                    return "❌ 未检测到打开的 PC 微信窗口。请先打开微信聊天窗口（不要最小化到托盘）。"

                msgs, mode_desc = self._extract_messages_dual_engine(wx_ctrl, self.target_hwnd)
                lines = [
                    f"✅ 成功定位微信窗口：{wx_ctrl.Name} (类名: {wx_ctrl.ClassName}, HWND: {self.target_hwnd})",
                    f"⚙️ 当前读取引擎：{mode_desc}",
                    f"📨 读取到底部最近消息数：{len(msgs)} 条",
                ]
                for i, cm in enumerate(msgs[-6:], 1):
                    side_tag = "[🙋自己/右]" if cm.is_self else "[👤对方/左]"
                    sdr_str = f"[{cm.sender}] " if cm.sender else ""
                    preview = cm.content.replace("\n", " / ")[:75]
                    lines.append(f"   {i}. {side_tag} {sdr_str}{preview}")
                if not msgs:
                    lines.append("⚠️ 未读取到消息气泡（请确保聊天窗口已打开、没有被最小化，且画面中有文字气泡）。")
                return "\n".join(lines)
        except Exception as e:
            return f"❌ 微信窗口诊断异常: {e}"

    def _log(self, text: str) -> None:
        ts = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        self.on_log(f"[{ts}] {text}")

    def _find_target_control(self, auto):
        if self.target_hwnd and user32.IsWindow(self.target_hwnd):
            try:
                ctrl = auto.ControlFromHandle(self.target_hwnd)
                if ctrl:
                    return ctrl
            except Exception:
                pass

        wx_windows = list_wechat_windows()
        if self.target_group_keyword:
            for w in wx_windows:
                if self.target_group_keyword in w.title:
                    self.target_hwnd = w.hwnd
                    return auto.ControlFromHandle(w.hwnd)

        if wx_windows:
            self.target_hwnd = wx_windows[0].hwnd
            return auto.ControlFromHandle(wx_windows[0].hwnd)

        return None

    # ------------------------------------------------------------------
    # 引擎 1：Windows UIAutomation 树读取（微信 3.x）
    # ------------------------------------------------------------------
    def _find_message_list_control(self, wx_ctrl):
        msg_list = wx_ctrl.ListControl(Name="消息", searchDepth=12)
        if msg_list.Exists(0, 0):
            return msg_list

        candidates = []
        for ctrl, _ in _iter_controls_bfs(wx_ctrl, max_depth=10, max_nodes=400):
            cls = ctrl.ClassName or ""
            ct = ctrl.ControlTypeName
            if ct == "ListControl" or "ChatMessagePage" in cls or "MessageList" in cls:
                cname = (ctrl.Name or "").strip()
                if cname in ("会话", "联系人", "搜索结果", "通讯录"):
                    continue
                try:
                    rect = ctrl.BoundingRectangle
                    if rect.width() > 150 and rect.height() > 150:
                        score = (1000 if "消息" in cname else 0) + rect.left + rect.width()
                        candidates.append((score, ctrl))
                except Exception:
                    pass
        if candidates:
            candidates.sort(key=lambda x: x[0], reverse=True)
            return candidates[0][1]
        return None

    def _find_chat_input_control(self, wx_ctrl):
        win_rect = wx_ctrl.BoundingRectangle
        mid_y = win_rect.top + win_rect.height() * 0.5
        candidates = []
        for ctrl, _ in _iter_controls_bfs(wx_ctrl, max_depth=12, max_nodes=400):
            if ctrl.ControlTypeName in ("EditControl", "DocumentControl"):
                cname = (ctrl.Name or "").strip()
                cls = ctrl.ClassName or ""
                if "搜索" in cname or "Search" in cls:
                    continue
                try:
                    r = ctrl.BoundingRectangle
                    if r.top >= mid_y and r.width() > 80:
                        score = (1000 if cname in ("输入", "微信输入") else 0) + r.top
                        candidates.append((score, ctrl))
                except Exception:
                    pass
        if candidates:
            candidates.sort(key=lambda x: x[0], reverse=True)
            return candidates[0][1]
        return None

    def _detect_item_sender_and_side_uia(self, item, list_rect, raw_name: str) -> Tuple[str, bool, str]:
        sender = ""
        extracted_texts: List[str] = []
        avatar_center_x: Optional[float] = None
        bubble_rect = None

        list_left, list_right = list_rect.left, list_rect.right
        list_width = max(1, list_right - list_left)
        list_mid_x = (list_left + list_right) / 2.0

        try:
            for sub, depth in _iter_controls_bfs(item, max_depth=5, max_nodes=45):
                if depth == 0:
                    continue
                ct = sub.ControlTypeName
                cls = sub.ClassName or ""
                sname = (sub.Name or "").strip()

                if ct == "ButtonControl":
                    try:
                        r = sub.BoundingRectangle
                        if 18 <= r.width() <= 70 and 18 <= r.height() <= 70 and avatar_center_x is None:
                            avatar_center_x = (r.left + r.right) / 2.0
                    except Exception:
                        pass
                    if sname and sname != raw_name and not sender:
                        sender = sname

                if "ChatBubble" in cls or "ChatTextItemView" in cls or ct == "TextControl":
                    try:
                        r = sub.BoundingRectangle
                        if 10 < r.width() < list_width * 0.85 and r.height() > 10:
                            if bubble_rect is None or "ChatTextItemView" in cls or "ChatBubble" in cls:
                                bubble_rect = r
                    except Exception:
                        pass
                    if sname and sname not in extracted_texts and sname != sender:
                        extracted_texts.append(sname)
        except Exception:
            pass

        is_self = False
        if avatar_center_x is not None:
            if avatar_center_x > list_mid_x + 20:
                is_self = True
        elif bubble_rect is not None:
            left_gap = bubble_rect.left - list_left
            right_gap = list_right - bubble_rect.right
            if right_gap < left_gap - 25 and right_gap < 120:
                is_self = True

        content = raw_name
        if not content and extracted_texts:
            body_parts = [t for t in extracted_texts if t != sender]
            content = "\n".join(body_parts) if body_parts else "\n".join(extracted_texts)

        return sender, is_self, content.strip()

    def _extract_messages_uia(self, wx_ctrl) -> List[CapturedMessage]:
        items_out: List[CapturedMessage] = []
        try:
            msg_list = self._find_message_list_control(wx_ctrl)
            if not msg_list:
                return items_out

            list_rect = msg_list.BoundingRectangle
            children = msg_list.GetChildren()
            recent_children = children[-15:] if len(children) > 15 else children

            for item in recent_children:
                raw_name = (item.Name or "").strip()
                sender, is_self, content = self._detect_item_sender_and_side_uia(item, list_rect, raw_name)
                if not content:
                    continue
                if re.match(r"^(\d{1,2}:\d{2}|昨天.*|星期.*|周.*|\d+月\d+日.*|以下为新消息|查看更多消息)$", content):
                    continue
                if self._is_recently_sent_echo(content):
                    is_self = True
                fp = f"{int(is_self)}::{sender}::{content}"
                items_out.append(CapturedMessage(fingerprint=fp, sender=sender, content=content, is_self=is_self))
        except Exception:
            pass
        return items_out

    # ------------------------------------------------------------------
    # 引擎 2：DWM 后台抓取 + 视觉气泡分割 + Windows 原生 WinRT OCR（微信 4.0 专用）
    # ------------------------------------------------------------------
    def _extract_messages_visual_ocr(self, hwnd: int, title: str = "") -> Tuple[List[CapturedMessage], str]:
        items_out: List[CapturedMessage] = []
        if not self._ocr.is_available():
            return items_out, "系统OCR不可用"

        win_img, (win_left, win_top, _win_right, _win_bottom), cap_mode = capture_window_image(hwnd)
        if win_img is None:
            return items_out, cap_mode

        chat_x0, chat_y0, chat_x1, chat_y1, input_click_y = _locate_chat_pane_rect(win_img, title)
        self._cached_input_abs_xy = (
            win_left + (chat_x0 + chat_x1) // 2,
            win_top + input_click_y,
        )

        chat_img = win_img.crop((chat_x0, chat_y0, chat_x1, chat_y1))
        bubbles = detect_chat_bubbles_adaptive(chat_img)

        recent_bubbles = bubbles[-6:] if len(bubbles) > 6 else bubbles
        for y0, y1, x0, x1, is_self in recent_bubbles:
            msg_crop = chat_img.crop((x0, y0, x1, y1))
            # 采样气泡中心点底色做边缘留白填充，提高 OCR 首尾字符准确率
            pad_col = msg_crop.getpixel((min(msg_crop.width - 1, 6), min(msg_crop.height - 1, 6)))
            lines = self._ocr.recognize_text_lines(msg_crop, scale=3.0, pad_color=pad_col)
            content = _join_bubble_ocr_lines(lines, msg_crop.width)
            if not content:
                continue

            sender = ""
            if not is_self and y0 >= 18:
                sdr_y0 = max(0, y0 - 32)
                sdr_crop = chat_img.crop((max(0, x0 - 20), sdr_y0, min(chat_img.width, x0 + 250), y0))
                sdr_lines = self._ocr.recognize_text_lines(sdr_crop, scale=3.0, pad_color=(250, 250, 250))
                if sdr_lines:
                    sdr_cand = sdr_lines[-1][4].strip()
                    if not re.match(r"^(\d{1,2}[:：]\d{2}|星期.*|周.*|\d+月\d+日.*)$", sdr_cand):
                        sender = sdr_cand

            if self._is_recently_sent_echo(content):
                is_self = True

            fp = f"{int(is_self)}::{sender}::{content}"
            items_out.append(CapturedMessage(fingerprint=fp, sender=sender, content=content, is_self=is_self))

        if not items_out and chat_img.height > 120:
            bot_y = int(chat_img.height * 0.45)
            bot_crop = chat_img.crop((0, bot_y, chat_img.width, chat_img.height))
            lines = self._ocr.recognize_text_lines(bot_crop, scale=2.5, pad_color=(250, 250, 250))
            mid_x = chat_img.width * 0.48
            for _y0, lx0, lx1, _y1, txt in lines[-8:]:
                if re.match(r"^(\d{1,2}[:：]\d{2}|星期.*|周.*|\d+月\d+日.*|.*拍了拍.*|.*撤回了一条消息.*)$", txt):
                    continue
                cx = (lx0 + lx1) / 2.0
                is_self = (lx0 > mid_x) or (cx > chat_img.width * 0.62 and (chat_img.width - lx1) < lx0)
                if self._is_recently_sent_echo(txt):
                    is_self = True
                fp = f"{int(is_self)}::::{txt}"
                items_out.append(CapturedMessage(fingerprint=fp, sender="", content=txt, is_self=is_self))

        return items_out, cap_mode

    def _is_recently_sent_echo(self, content: str) -> bool:
        now_ts = time.time()
        self._recent_sent_replies = [(txt, ts) for txt, ts in self._recent_sent_replies if now_ts - ts < 30.0]
        nc = _norm_msg_for_match(content)
        return any(nc == _norm_msg_for_match(txt) for txt, _ in self._recent_sent_replies)

    def _extract_messages_dual_engine(self, wx_ctrl, hwnd: int) -> Tuple[List[CapturedMessage], str]:
        uia_msgs = self._extract_messages_uia(wx_ctrl)
        if uia_msgs:
            return uia_msgs, "UIAutomation 原生控件树 (微信3.x/无障碍模式)"

        ocr_msgs, cap_mode = self._extract_messages_visual_ocr(hwnd, title=(wx_ctrl.Name or ""))
        if self._ocr.is_available():
            return ocr_msgs, f"视觉气泡 + Windows原生OCR ({cap_mode})"
        return [], "UIAutomation未读取到控件且系统OCR不可用"

    # ------------------------------------------------------------------
    # 发送回复与主监听循环
    # ------------------------------------------------------------------
    def send_reply_to_wechat(self, reply_text: str) -> Tuple[bool, str]:
        now_ts = time.time()
        if now_ts - self._last_send_time < self.cooldown_sec:
            return False, f"处于防刷屏冷却期内 ({self.cooldown_sec}s)"

        self._recent_sent_replies.append((reply_text.strip(), now_ts))

        if not self.auto_send_enabled:
            self._last_send_time = now_ts
            return True, f"【演练模式/仅提醒】已模拟发送回复：「{reply_text}」（未实际向微信按键）"

        try:
            import uiautomation as auto
            wx_ctrl = self._find_target_control(auto)
            if not wx_ctrl or not self.target_hwnd:
                return False, "未找到打开的微信窗口，无法自动发送"

            if not set_clipboard_text(reply_text):
                return False, "写入系统剪贴板失败"

            _activate_window(self.target_hwnd)
            time.sleep(0.08)

            input_ctrl = self._find_chat_input_control(wx_ctrl)
            if input_ctrl:
                input_ctrl.SetFocus()
                time.sleep(0.03)
            elif self._cached_input_abs_xy is not None:
                cx, cy = self._cached_input_abs_xy
                auto.Click(cx, cy, waitTime=0.05)
            else:
                rect = wx_ctrl.BoundingRectangle
                if rect.width() > 100 and rect.height() > 100:
                    click_x = int(rect.left + rect.width() * 0.65)
                    click_y = int(rect.top + rect.height() * 0.88)
                    auto.Click(click_x, click_y, waitTime=0.05)

            auto.SendKeys("{Ctrl}v{Enter}", waitTime=0.03)
            self._last_send_time = time.time()
            return True, f"已向微信群极速发送：「{reply_text}」"
        except Exception as e:
            return False, f"微信自动发送异常: {e}"

    def _monitor_loop(self) -> None:
        try:
            import uiautomation as auto
            with auto.UIAutomationInitializerInThread():
                wx_ctrl = self._find_target_control(auto)
                if not wx_ctrl or not self.target_hwnd:
                    self._log("⚠️ 暂未检测到打开的微信窗口。请打开 PC 微信聊天窗口（可被遮挡，但不要最小化），监控将自动连接...")
                else:
                    init_msgs, mode_desc = self._extract_messages_dual_engine(wx_ctrl, self.target_hwnd)
                    self._last_engine_used = mode_desc
                    self._last_tail = init_msgs
                    self._log(f"✅ 已连接微信窗口: {wx_ctrl.Name} (Class={wx_ctrl.ClassName})")
                    self._log(f"⚙️ 启用读取引擎: {mode_desc}")
                    self._log(f"📌 已记录当前聊天底部 {len(init_msgs)} 条历史消息基线，仅对后续新发出的消息进行判断。")

                warned_no_window = (wx_ctrl is None or self.target_hwnd is None)
                baseline_initialized = not warned_no_window

                while not self._stop_event.is_set():
                    try:
                        wx_ctrl = self._find_target_control(auto)
                        if not wx_ctrl or not self.target_hwnd or user32.IsIconic(self.target_hwnd):
                            if not warned_no_window:
                                self._log("⚠️ 微信窗口未打开或已最小化，请保持聊天窗口处于非最小化状态（可被其他窗口遮挡）...")
                                warned_no_window = True
                            time.sleep(0.8)
                            continue

                        if warned_no_window:
                            self._log(f"✅ 已重新连接微信窗口: {wx_ctrl.Name}")
                            warned_no_window = False

                        recent, mode_desc = self._extract_messages_dual_engine(wx_ctrl, self.target_hwnd)
                        if mode_desc != self._last_engine_used and recent:
                            self._last_engine_used = mode_desc
                            self._log(f"⚙️ 切换读取引擎: {mode_desc}")

                        if not baseline_initialized:
                            if recent:
                                self._last_tail = recent
                                self._log(f"📌 已记录 {len(recent)} 条历史消息基线，开始实时监听新消息！")
                                baseline_initialized = True
                            time.sleep(self.poll_interval_sec)
                            continue

                        if recent:
                            new_items = _compute_new_tail_items(self._last_tail, recent)
                            self._last_tail = recent
                            now_ts = time.time()
                            self._recent_processed = {k: t for k, t in self._recent_processed.items() if now_ts - t < 180.0}
                            for cm in new_items:
                                norm_k = f"{int(cm.is_self)}::{_norm_msg_for_match(cm.content)}"
                                if norm_k in self._recent_processed:
                                    continue
                                self._recent_processed[norm_k] = now_ts
                                self.on_new_message(cm.sender, cm.content, cm.is_self)
                    except Exception as loop_err:
                        self._log(f"⚠️ 单次轮询遇到临时异常（监控继续运行）: {loop_err}")
                        time.sleep(0.5)

                    time.sleep(self.poll_interval_sec)
        except Exception as e:
            self._log(f"❌ 监控线程异常退出: {e}")
        finally:
            self.is_running = False
            if self.on_status_change:
                self.on_status_change(False, "监控已停止")
