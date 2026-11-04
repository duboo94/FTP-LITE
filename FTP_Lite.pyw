r"""
FTP Lite - FANUC 장비 FTP 데이터 서버 파일 전송 (간단 버전)

- 설비마다 FTP 로 접속해서 파일을 올리고 내려받는 창 하나만 있다 (모니터링 · 알람 · 휴대폰 연결 · 라이선스 없음)
- 위쪽: 설비의 데이터 서버 (가로로 나란히)  /  아래쪽: 내 PC 폴더  /  끌어서 놓기로 전송
- 시계 옆 트레이 아이콘 + 단축키(기본 Ctrl+Alt+F)로 보이기/숨기기, 전송이 끝나면 자동으로 숨김
- 설정 파일: %LOCALAPPDATA%\\FTP Lite\\Settings.json  (설비 목록 + 설정, 비밀번호는 암호화하지 않음)

실행: 파일을 더블클릭 (.pyw 라서 콘솔 창 없이 실행, 표준 라이브러리만 사용)
"""
import ctypes
import json
import os
import queue
import re
import shutil
import socket
import struct
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
from dataclasses import dataclass
from ftplib import FTP, all_errors, error_perm, error_temp
from pathlib import Path
from typing import Optional
from tkinter import filedialog, font as tkfont, messagebox, ttk

# --------------------------------------------------------------------------
# 저장 파일 (하나): 내 문서\FTP Manager\Settings.json
#   {"sites": {이름: {host, user, password}}, "settings": {창 위치·크기, 투명도, 단축키 ...}}
#   폴더 지정: %LOCALAPPDATA%\FTP Manager\data_root.txt > 내 문서\FTP Manager > %LOCALAPPDATA%\FTP Manager. 파일이 없으면 처음 실행으로 보고 설정 창을 먼저 연다.
# --------------------------------------------------------------------------
FTP_PORT = 21
DATA_NAME = "Settings.json"


def _app_dir():
    here = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
    try:
        probe = here / ".write_test"
        probe.write_text("", encoding="utf-8")
        probe.unlink()
        return here
    except OSError:                 # 프로그램 폴더에 쓸 수 없으면 사용자 폴더에 저장
        return Path.home()


APP_FOLDER = "FTP Lite"


def _documents_dir():
    """내 문서 폴더 (M-Link 와 같은 방식: 윈도우 '알려진 폴더' 조회, 실패하면 USERPROFILE\\Documents)."""
    if sys.platform == "win32":
        try:
            import ctypes
            from ctypes import wintypes

            class GUID(ctypes.Structure):
                _fields_ = [("Data1", ctypes.c_ulong), ("Data2", ctypes.c_ushort),
                            ("Data3", ctypes.c_ushort), ("Data4", ctypes.c_ubyte * 8)]
            guid = GUID()
            ctypes.windll.ole32.CLSIDFromString("{FDD39AD0-238F-46AF-ADB4-6C85480369C7}", ctypes.byref(guid))
            out = ctypes.c_wchar_p()
            ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(guid), 0, None, ctypes.byref(out))
            path = out.value
            ctypes.windll.ole32.CoTaskMemFree(out)
            if path:
                return Path(path)
        except Exception:
            pass
        return Path(os.environ.get("USERPROFILE") or str(Path.home())) / "Documents"
    return Path.home() / "Documents"


def _local_dir():
    """%LOCALAPPDATA%\\FTP_Manager : 이 PC 전용 자리 (폴더 지정 파일, 문서 폴더에 못 쓸 때의 대체 저장소)."""
    base = os.environ.get("LOCALAPPDATA")
    base = Path(base) if base else Path.home() / "AppData" / "Local"
    return base / APP_FOLDER


def _can_write_dir(d):
    try:
        d.mkdir(parents=True, exist_ok=True)
        probe = d / (".nc_write_probe_%d" % os.getpid())
        probe.write_bytes(b"")
        probe.unlink()
        return True
    except OSError:
        return False


def _data_root():
    """데이터 폴더: %LOCALAPPDATA%\\FTP Manager\\data_root.txt 에 경로가 있으면 그 폴더,
    아니면 %LOCALAPPDATA%\\FTP Manager (EXE 든 .pyw 든 항상 같은 자리 - 보호 폴더 문제 없음).
    이 폴더에 설정이 아직 없고 예전 위치(내 문서\\FTP Manager)에 있으면 한 번 복사해 온다."""
    pointer = _local_dir() / "data_root.txt"
    try:
        text = pointer.read_text(encoding="utf-8").strip()
        if text and _can_write_dir(Path(text)):
            return Path(text)
    except (OSError, ValueError):
        pass
    root = _local_dir()
    if _can_write_dir(root):
        return root
    root = _documents_dir() / APP_FOLDER
    if _can_write_dir(root):
        return root
    return Path.home()


def _pick_data_path():
    return _data_root() / DATA_NAME


APP_DIR = _app_dir()
DATA_PATH = _pick_data_path()
SITES_PATH = APP_CONFIG_PATH = DATA_PATH
FIRST_RUN = not DATA_PATH.exists()          # 설정 파일이 없는 첫 실행 (이후 기본값을 파일에 쓰기 전에 먼저 확인)


def _read_json(path, old_path=None):
    for p in (path, old_path):
        if p is None:
            continue
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
        except (OSError, ValueError):
            continue
    return {}


def _write_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)               # 쓰다가 꺼져도 파일이 깨지지 않도록


def _load_all():
    data = _read_json(DATA_PATH)
    if isinstance(data.get("sites"), dict) or isinstance(data.get("settings"), dict):
        return {"sites": data["sites"] if isinstance(data.get("sites"), dict) else {},
                "settings": data["settings"] if isinstance(data.get("settings"), dict) else {}}
    return {"sites": {}, "settings": {}}


def clean_sites(raw):
    """{이름: {"host", "user", "password", "ftp"}} (순서 유지, 형식이 이상한 항목은 버림)."""
    sites = {}
    if isinstance(raw, dict):
        for name, s in raw.items():
            if isinstance(s, dict):
                sites[str(name)] = {"host": str(s.get("host", "")).strip(),
                                    "user": str(s.get("user", "")),
                                    "password": str(s.get("password", "")),
                                    "ftp": True}
    return sites


def load_sites():
    return clean_sites(_load_all()["sites"])


def save_sites(sites):
    data = _load_all()
    data["sites"] = sites
    _write_json(DATA_PATH, data)


EDITOR_PATH = [""]                 # 오른쪽 클릭으로 파일을 열 프로그램 (비어 있으면 메모장). 설정 > 기타 에서 정한다


def load_app_cfg():
    cfg = dict(_load_all()["settings"])
    EDITOR_PATH[0] = str(cfg.get("editor_path", "") or "").strip()
    return cfg


def save_app_cfg(cfg):
    try:
        data = _load_all()
        data["settings"] = cfg
        _write_json(DATA_PATH, data)
    except OSError:
        pass


# --------------------------------------------------------------------------
# 장비 찾기: 같은 네트워크에서 21번(FTP) 포트가 열린 IP 를 찾는다
# --------------------------------------------------------------------------
def local_prefixes(extra_hosts=()):
    """이 PC 의 IPv4 주소와 등록된 호스트에서 '192.168.123' 같은 앞 3자리 목록을 만든다."""
    ips = set()
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ips.add(info[4][0])
    except OSError:
        pass
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:   # 실제로 보내지 않음
            s.connect(("10.255.255.255", 1))
            ips.add(s.getsockname()[0])
    except OSError:
        pass
    ips.update(extra_hosts)
    out = []
    for ip in ips:
        parts = ip.split(".")
        if len(parts) == 4 and all(p.isdigit() for p in parts) and not ip.startswith(("127.", "169.254.", "0.")):
            prefix = ".".join(parts[:3])
            if prefix not in out:
                out.append(prefix)
    return out


def valid_prefix(text):
    parts = text.strip().split(".")
    return len(parts) == 3 and all(p.isdigit() and 0 <= int(p) <= 255 for p in parts)


def probe_host(ip, timeout=0.6):
    """21번 포트가 열려 있으면 인사말(빈 문자열일 수 있음)을 (인사말,) 로, 아니면 None."""
    try:
        sock = socket.create_connection((ip, FTP_PORT), timeout=timeout)
    except OSError:
        return None
    banner = ""
    try:
        sock.settimeout(1.0)
        banner = sock.recv(256).decode("latin-1", "replace").strip().splitlines()[0]
    except (OSError, IndexError):
        pass
    finally:
        try:
            sock.sendall(b"QUIT\r\n")
        except OSError:
            pass
        sock.close()
    return (banner,)


# --------------------------------------------------------------------------
# 애플 스타일 UI 부품: 밝은 회색 배경, 흰 카드, 파란 강조, 둥근 모서리
# --------------------------------------------------------------------------
C_WIN = "#f2f2f7"          # 창 배경
C_CARD = "#ffffff"         # 카드 배경
C_LINE = "#e5e5ea"         # 구분선
C_BORDER = "#d1d1d6"       # 입력칸 테두리
C_TEXT = "#1d1d1f"         # 기본 글자
C_SUB = "#6e6e73"          # 보조 글자
C_HINT = "#75757d"         # 힌트 글자
C_BLUE = "#0071e3"         # 강조(파랑)
C_BLUE_HOVER = "#0077ed"
C_BLUE_PRESS = "#006edb"
C_GRAY_BTN = "#e8e8ed"
C_GRAY_HOVER = "#dedee3"
C_GRAY_PRESS = "#d2d2d7"
C_GREEN = "#34c759"
C_RED = "#d70015"

UI_FONT = "Malgun Gothic"


def pick_ui_font(root):
    """설치된 글꼴 중 한글이 깨끗한 것을 고른다."""
    global UI_FONT
    try:
        families = set(tkfont.families(root))
    except tk.TclError:
        return
    for name in ("Malgun Gothic", "맑은 고딕", "Segoe UI", "Noto Sans CJK KR", "NanumGothic", "DejaVu Sans"):
        if name in families:
            UI_FONT = name
            return


def ui_font(size=10, bold=False):
    return (UI_FONT, size, "bold") if bold else (UI_FONT, size)


def rr_points(x1, y1, x2, y2, r):
    """둥근 사각형 (smooth 다각형) 좌표."""
    r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    return [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2, x2 - r, y2,
            x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]


def rr_points_sel(x1, y1, x2, y2, tl, tr, br, bl):
    """모서리마다 반지름을 따로 주는 둥근 사각형 (0 이면 각진 모서리)."""
    lim = min((x2 - x1) / 2, (y2 - y1) / 2)
    tl, tr, br, bl = (max(0, min(r, lim)) for r in (tl, tr, br, bl))
    return [x1 + tl, y1, x2 - tr, y1, x2, y1, x2, y1 + tr, x2, y2 - br, x2, y2, x2 - br, y2,
            x1 + bl, y2, x1, y2, x1, y2 - bl, x1, y1 + tl, x1, y1]


def blend(c1, c2, t):
    """c1 과 c2 를 t(0~1) 비율로 섞은 색."""
    a = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02x%02x%02x" % tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


def apply_styles(root):
    """ttk 위젯(목록)을 애플 스타일로."""
    st = ttk.Style(root)
    try:
        st.theme_use("clam")
    except tk.TclError:
        pass
    st.configure("Apple.Treeview", background=C_CARD, fieldbackground=C_CARD, foreground=C_TEXT,
                 borderwidth=0, relief="flat", rowheight=32, font=ui_font(10))
    st.map("Apple.Treeview", background=[("selected", C_BLUE)], foreground=[("selected", "#ffffff")])
    st.layout("Apple.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
    st.configure("Apple.Treeview.Heading", background=C_CARD, foreground=C_SUB, relief="flat",
                 borderwidth=0, font=ui_font(9, True), padding=(8, 6))
    st.map("Apple.Treeview.Heading", background=[("active", C_CARD)])


class RoundButton(tk.Canvas):
    """둥근 버튼. kind: primary(파랑) / secondary(회색) / danger(빨간 글자) / plain(글자만)."""

    KINDS = {
        "primary": (C_BLUE, C_BLUE_HOVER, C_BLUE_PRESS, "#ffffff"),
        "secondary": (C_GRAY_BTN, C_GRAY_HOVER, C_GRAY_PRESS, C_TEXT),
        "danger": (C_GRAY_BTN, C_GRAY_HOVER, C_GRAY_PRESS, C_RED),
        "plain": (None, C_WIN, C_LINE, C_BLUE),
    }

    def __init__(self, master, text, command=None, kind="secondary", width=None, height=32,
                 radius=9, outer=C_CARD, size=10):
        self._font = tkfont.Font(family=UI_FONT, size=size, weight="bold" if kind == "primary" else "normal")
        self._text, self._command, self._kind = text, command, kind
        self._fixed_w, self._h, self._r = width, height, radius
        self._enabled, self._hover, self._down = True, False, False
        self._bw = self._calc_w()
        super().__init__(master, width=self._bw, height=height, bg=outer, highlightthickness=0, bd=0,
                         cursor="hand2")
        self.bind("<Enter>", lambda _e: self._state(hover=True))
        self.bind("<Leave>", lambda _e: self._state(hover=False, down=False))
        self.bind("<ButtonPress-1>", lambda _e: self._state(down=True))
        self.bind("<ButtonRelease-1>", self._release)
        self._draw()

    def _calc_w(self):
        return self._fixed_w or self._font.measure(self._text) + 30

    def _state(self, hover=None, down=None):
        if hover is not None:
            self._hover = hover
        if down is not None:
            self._down = down
        self._draw()

    def _release(self, e):
        fire = self._down and self._enabled and 0 <= e.x <= self._bw and 0 <= e.y <= self._h
        self._down = False
        self._draw()
        if fire and self._command:
            self._command()

    def _draw(self):
        try:
            self.delete("all")
            base, hov, prs, fg = self.KINDS[self._kind]
            if not self._enabled:
                fill, fg = (C_WIN if base else None), "#c7c7cc"
            elif self._down and self._hover:
                fill = prs
            elif self._hover:
                fill = hov
            else:
                fill = base
            if fill:
                self.create_polygon(rr_points(1, 1, self._bw - 1, self._h - 1, self._r), smooth=True,
                                    fill=fill, outline="")
            self.create_text(self._bw / 2, self._h / 2, text=self._text, font=self._font, fill=fg)
        except tk.TclError:
            pass

    def configure(self, cnf=None, **kw):
        redraw = False
        if "state" in kw:
            self._enabled = kw.pop("state") != "disabled"
            redraw = True
        if "text" in kw:
            self._text = kw.pop("text")
            self._bw = self._calc_w()
            kw["width"] = self._bw
            redraw = True
        if kw or cnf:
            super().configure(cnf, **kw)
        if redraw:
            self._draw()

    config = configure


_SWITCH_IMG = {}


def _switch_image(on, bg):
    """안티앨리어싱(4x4 표본) 된 켜기/끄기 스위치 그림."""
    key = (on, bg)
    if key in _SWITCH_IMG:
        return _SWITCH_IMG[key]
    W, H, SS = 46, 28, 4
    track = tuple(int((C_GREEN if on else "#d1d1d6")[i:i + 2], 16) for i in (1, 3, 5))
    back = tuple(int(bg[i:i + 2], 16) for i in (1, 3, 5))
    edge = (0xd9, 0xd9, 0xde)
    kx = 32.0 if on else 14.0
    rows = []
    for y in range(H):
        row = []
        for x in range(W):
            acc = [0, 0, 0]
            for sy in range(SS):
                for sx in range(SS):
                    px, py = x + (sx + .5) / SS, y + (sy + .5) / SS
                    if (px - kx) ** 2 + (py - 14) ** 2 <= 11 ** 2:
                        c = (255, 255, 255)
                    elif (px - kx) ** 2 + (py - 14) ** 2 <= 11.9 ** 2:
                        c = edge
                    else:
                        qx = min(max(px, 14.0), 32.0)              # 알약 모양 = 선분 주변 반지름 13
                        c = track if (px - qx) ** 2 + (py - 14) ** 2 <= 13 ** 2 else back
                    acc[0] += c[0]
                    acc[1] += c[1]
                    acc[2] += c[2]
            n = SS * SS
            row.append("#%02x%02x%02x" % (acc[0] // n, acc[1] // n, acc[2] // n))
        rows.append("{" + " ".join(row) + "}")
    img = tk.PhotoImage(width=W, height=H)
    img.put(" ".join(rows))
    _SWITCH_IMG[key] = img
    return img


class Switch(tk.Label):
    """iOS 스타일 켜기/끄기 스위치 (BooleanVar 와 연결)."""

    def __init__(self, master, variable, outer=C_CARD):
        super().__init__(master, bg=outer, bd=0, highlightthickness=0, cursor="hand2")
        self._outer = outer
        self.var = variable
        self.bind("<Button-1>", lambda _e: self.var.set(not self.var.get()))
        self.var.trace_add("write", lambda *_: self._draw())
        self._draw()

    def _draw(self):
        try:
            self.configure(image=_switch_image(bool(self.var.get()), self._outer))
        except tk.TclError:
            pass


class Slider(tk.Canvas):
    """가로 슬라이더 (IntVar 와 연결). 값이 바뀔 때마다 command(값)."""

    def __init__(self, master, variable, lo, hi, command=None, width=220, outer=C_CARD):
        super().__init__(master, width=width, height=28, bg=outer, highlightthickness=0, bd=0, cursor="hand2")
        self.var, self.lo, self.hi, self._cmd, self._bw = variable, lo, hi, command, width
        self.bind("<Button-1>", self._set)
        self.bind("<B1-Motion>", self._set)
        self.var.trace_add("write", lambda *_: self._draw())
        self._draw()

    def _set(self, e):
        span = self._bw - 24
        v = self.lo + (min(max(e.x, 12), self._bw - 12) - 12) / span * (self.hi - self.lo)
        v = int(round(v))
        if v != self.var.get():
            self.var.set(v)
        if self._cmd:
            self._cmd(v)

    def _draw(self):
        try:
            self.delete("all")
            span = self._bw - 24
            x = 12 + (self.var.get() - self.lo) / (self.hi - self.lo) * span
            self.create_line(12, 14, self._bw - 12, 14, width=4, fill="#d1d1d6", capstyle="round")
            self.create_line(12, 14, x, 14, width=4, fill=C_BLUE, capstyle="round")
            self.create_oval(x - 10, 4, x + 10, 24, fill="#ffffff", outline="#c7c7cc")
        except (tk.TclError, ValueError):
            pass


class Segmented(tk.Canvas):
    """분할 탭 (애플 세그먼트 컨트롤)."""

    def __init__(self, master, labels, command, width=300, height=34, outer=C_WIN):
        super().__init__(master, width=width, height=height, bg=outer, highlightthickness=0, bd=0, cursor="hand2")
        self.labels, self._cmd, self._bw, self._h, self.selected = labels, command, width, height, 0
        self._font = tkfont.Font(family=UI_FONT, size=10)
        self._bold = tkfont.Font(family=UI_FONT, size=10, weight="bold")
        self.bind("<Button-1>", self._click)
        self._draw()

    def _click(self, e):
        idx = min(len(self.labels) - 1, max(0, int(e.x // (self._bw / len(self.labels)))))
        if idx != self.selected:
            self.select(idx)
            self._cmd(idx)

    def select(self, idx):
        self.selected = idx
        self._draw()

    def _draw(self):
        self.delete("all")
        n = len(self.labels)
        seg = (self._bw - 4) / n
        self.create_polygon(rr_points(0, 0, self._bw, self._h, 10), smooth=True, fill="#e3e3e8", outline="")
        for i, text in enumerate(self.labels):
            x1 = 2 + i * seg
            if i == self.selected:
                self.create_polygon(rr_points(x1, 2, x1 + seg, self._h - 2, 8), smooth=True,
                                    fill="#ffffff", outline="#d9d9de")
            self.create_text(x1 + seg / 2, self._h / 2, text=text,
                             font=self._bold if i == self.selected else self._font,
                             fill=C_TEXT if i == self.selected else C_SUB)


class RoundedCard(tk.Canvas):
    """둥근 흰 카드. 내용은 self.inner 프레임에 넣는다."""

    def __init__(self, master, outer=C_WIN, fill=C_CARD, radius=12, pad=14, expand=False, width=None):
        super().__init__(master, bg=outer, highlightthickness=0, bd=0, height=40, width=width or 100)
        self.fill, self.radius, self.pad, self.expand = fill, radius, pad, expand
        self.inner = tk.Frame(self, bg=fill)
        self._win = self.create_window(pad, pad, window=self.inner, anchor="nw")
        self.bind("<Configure>", self._on_canvas)
        self.inner.bind("<Configure>", self._on_inner)

    def _on_inner(self, _e):
        want = self.inner.winfo_reqheight() + 2 * self.pad
        if int(float(self.cget("height"))) != want:
            self.configure(height=want)

    def _on_canvas(self, e):
        self.delete("bg")
        self.create_polygon(rr_points(0, 0, e.width - 1, e.height - 1, self.radius), smooth=True,
                            fill=self.fill, outline=C_LINE, tags="bg")
        self.tag_lower("bg")
        opts = {"width": max(1, e.width - 2 * self.pad)}
        if self.expand:
            opts["height"] = max(1, e.height - 2 * self.pad)
        self.itemconfigure(self._win, **opts)


def section(parent, title=None, expand=False, pad=14, outer=C_WIN, width=None):
    """제목(작은 회색 글자) + 둥근 카드. (바깥 프레임, 카드 안쪽 프레임)을 돌려준다."""
    wrap = tk.Frame(parent, bg=outer)
    if title:
        tk.Label(wrap, text=title, font=ui_font(9, True), bg=outer, fg=C_SUB, anchor="w").pack(
            fill="x", padx=6, pady=(0, 5))
    card = RoundedCard(wrap, outer=outer, pad=pad, expand=expand, width=width)
    card.pack(fill="both" if expand else "x", expand=expand)
    return wrap, card.inner


def fsection(parent, title=None, expand=False, width=None, **_):
    """상자 없는 단순 구역: 작은 굵은 제목 + 내용. (바깥 프레임, 안쪽 프레임)."""
    wrap = tk.Frame(parent, bg=C_CARD)
    if title:
        tk.Label(wrap, text=title, font=ui_font(9, True), bg=C_CARD, fg=C_SUB, anchor="w").pack(fill="x", pady=(0, 2))
    inner = tk.Frame(wrap, bg=C_CARD, width=width or 0)
    inner.pack(fill="both" if expand else "x", expand=expand)
    return wrap, inner


class ScrollPage(tk.Frame):
    """내용이 창보다 길면 스크롤바와 마우스 휠로 넘겨 볼 수 있는 쪽. 내용은 self.inner 에 넣는다."""

    def __init__(self, parent):
        super().__init__(parent, bg=C_CARD)
        self.canvas = tk.Canvas(self, bg=C_CARD, highlightthickness=0, bd=0)
        self.bar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview, style="ML.Vertical.TScrollbar")
        self.canvas.configure(yscrollcommand=self._on_scroll)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(self.canvas, bg=C_CARD)
        self._win = self.canvas.create_window(0, 0, window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda _e: self._fit())
        self.canvas.bind("<Configure>", lambda e: (self.canvas.itemconfigure(self._win, width=e.width), self._fit()))
        self.bind("<Enter>", self._bind_wheel)
        self.bind("<Leave>", self._unbind_wheel)

    def _fit(self):
        self.canvas.configure(scrollregion=(0, 0, self.inner.winfo_reqwidth(), self.inner.winfo_reqheight()))
        self._on_scroll(*self.canvas.yview())

    def _on_scroll(self, lo, hi):
        lo, hi = float(lo), float(hi)
        self.bar.set(lo, hi)
        if lo <= 0.0 and hi >= 1.0:
            self.bar.pack_forget()                                   # 한 화면에 다 들어가면 스크롤바를 숨긴다
        elif not self.bar.winfo_ismapped():
            self.bar.pack(side="right", fill="y", padx=(6, 0))

    def _scrollable(self):
        lo, hi = self.canvas.yview()
        return not (lo <= 0.0 and hi >= 1.0)

    def _wheel(self, e):
        if not self._scrollable():
            return
        step = -1 if (getattr(e, "delta", 0) > 0 or getattr(e, "num", 0) == 4) else 1
        self.canvas.yview_scroll(step * 3, "units")

    def _bind_wheel(self, _e=None):
        for ev in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.bind_all(ev, self._wheel)

    def _unbind_wheel(self, e=None):
        if e is not None:                                            # 쪽 안의 다른 위젯으로 옮겨 갈 때는 유지
            try:
                w = self.winfo_containing(e.x_root, e.y_root)
            except (KeyError, tk.TclError):
                w = None
            while w is not None:
                if w is self:
                    return
                w = getattr(w, "master", None)
        for ev in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.unbind_all(ev)


def field_entry(parent, var, width=None, show=None, justify=None):
    """애플 스타일 입력칸 (파란 포커스 링)."""
    kw = {"width": width, "justify": "center"} if width else {}
    if justify:
        kw["justify"] = justify
    if show:
        kw["show"] = show
    return tk.Entry(parent, textvariable=var, font=ui_font(10), relief="flat", bd=0, bg=C_CARD, fg=C_TEXT,
                    insertbackground=C_TEXT, highlightthickness=1, highlightbackground=C_BORDER,
                    highlightcolor=C_BLUE, disabledbackground=C_WIN, disabledforeground=C_HINT,
                    selectbackground="#b3d7ff", selectforeground=C_TEXT, **kw)


class DeviceFinder(tk.Toplevel):
    """장비 찾기 창. 체크한 장비를 '추가'하면 on_pick([ip, ...]) 를 부른다."""

    def __init__(self, master, prefixes, registered, on_pick):
        super().__init__(master)
        self.title("장비 찾기")
        self.configure(bg=C_WIN)
        self.geometry("660x470")
        self.minsize(560, 380)
        self.transient(master)
        self.registered = registered          # 함수: ip -> 등록된 이름 또는 None
        self.on_pick = on_pick
        self.stop = threading.Event()
        self.results = queue.Queue()
        self.scanning = False
        self.checked = 0
        self.prefixes = prefixes
        self.total = 0
        self.status_var = tk.StringVar()

        head = tk.Frame(self, bg=C_WIN)
        head.pack(fill="x", padx=20, pady=(16, 10))
        tk.Label(head, text="장비 찾기", font=ui_font(18, True), bg=C_WIN, fg=C_TEXT).pack(anchor="w")
        tk.Label(head, textvariable=self.status_var, font=ui_font(9), bg=C_WIN, fg=C_SUB,
                 anchor="w").pack(fill="x", pady=(3, 0))

        foot = tk.Frame(self, bg=C_WIN)
        foot.pack(side="bottom", fill="x", padx=20, pady=(10, 16))
        RoundButton(foot, "닫기", self.close, outer=C_WIN).pack(side="right")
        RoundButton(foot, "추가", self.add, kind="primary", width=90, outer=C_WIN).pack(side="right", padx=8)
        tk.Label(foot, text="추가할 장비를 클릭해 선택하세요 (여러 대 가능, Ctrl+A = 전체)", font=ui_font(9), bg=C_WIN, fg=C_HINT).pack(side="left")
        self.picked = set()

        wrap, inner = section(self, None, expand=True, pad=6)
        wrap.pack(fill="both", expand=True, padx=20)
        cols = ("ip", "name", "banner")
        self.tree = ttk.Treeview(inner, columns=cols, show="headings", selectmode="none",
                                 style="Apple.Treeview")
        for col, text, width, anchor in (("ip", "IP 주소", 150, "w"),
                                         ("name", "등록된 이름", 160, "w"), ("banner", "FTP 응답", 340, "w")):
            self.tree.heading(col, text=text, anchor=anchor)
            self.tree.column(col, width=width, anchor=anchor)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Button-1>", self.on_click)
        self.bind("<Control-a>", self.select_all)
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.bind("<Escape>", lambda _e: self.close())
        self.start_scan()

    def start_scan(self):
        """이 PC 가 연결된 네트워크(들)의 .1 ~ .254 를 모두 찾는다."""
        prefixes = [p for p in self.prefixes if valid_prefix(p)]
        if not prefixes:
            self.status_var.set("이 PC 의 네트워크 주소를 찾지 못했습니다. 랜선 연결을 확인하세요.")
            return
        self.total = 254 * len(prefixes)
        self.stop.set()
        self.stop = threading.Event()
        self.results = queue.Queue()
        self.checked = 0
        self.scanning = True
        self.tree.delete(*self.tree.get_children())
        threading.Thread(target=self._scan, args=(prefixes, self.stop, self.results), daemon=True).start()
        if not getattr(self, "_polling", False):
            self._polling = True
            self._poll()

    @staticmethod
    def _scan(prefixes, stop, results):
        from concurrent.futures import ThreadPoolExecutor
        ips = [f"{prefix}.{i}" for prefix in prefixes for i in range(1, 255)]

        def one(ip):
            if stop.is_set():
                return
            found = probe_host(ip)
            results.put(("found", ip, found) if found else ("miss", ip, None))

        with ThreadPoolExecutor(max_workers=64) as pool:
            list(pool.map(one, ips))
        if not stop.is_set():
            results.put(("done", None, None))

    def _poll(self):
        if not self.winfo_exists():
            return
        try:
            while True:
                kind, ip, info = self.results.get_nowait()
                if kind == "found":
                    banner = info[0]
                    self.tree.insert("", "end", iid=ip, values=(ip, self.registered(ip) or "", banner))
                    if ip in self.picked:
                        self.tree.selection_add(ip)
                    self._sort()
                if kind == "done":
                    self.scanning = False
                else:
                    self.checked += 1
        except queue.Empty:
            pass
        found = len(self.tree.get_children())
        if self.scanning:
            ranges = ", ".join(f"{p}.x" for p in self.prefixes)
            self.status_var.set(f"{ranges} 찾는 중...  {self.checked}/{self.total}  ({found}대 찾음)")
        else:
            self.status_var.set(f"{found}대 찾음  ·  {len(self.picked)}대 선택" if found
                                else "21번 포트가 열린 장비를 찾지 못했습니다. 랜선·장비 전원을 확인하세요.")
        self.after(150, self._poll)

    def _sort(self):  # IP 끝자리 순
        items = sorted(self.tree.get_children(), key=lambda ip: int(ip.rsplit(".", 1)[1]))
        for i, iid in enumerate(items):
            self.tree.move(iid, "", i)

    def on_click(self, e):
        """줄을 누르면 설비 목록처럼 파란색으로 선택/해제 (여러 줄 가능)."""
        if self.tree.identify_region(e.x, e.y) != "heading":
            row = self.tree.identify_row(e.y)
            if row:
                if row in self.picked:
                    self.picked.discard(row)
                    self.tree.selection_remove(row)
                else:
                    self.picked.add(row)
                    self.tree.selection_add(row)
        return "break"

    def select_all(self, _e=None):
        ips = list(self.tree.get_children())
        self.picked = set(ips)
        self.tree.selection_set(ips)
        return "break"

    def add(self):
        ips = [ip for ip in self.tree.get_children() if ip in self.picked]
        if not ips:
            messagebox.showinfo("장비 찾기", "추가할 장비를 선택하세요.", parent=self)
            return
        if self.on_pick(ips) is False:
            return
        self.close()

    def close(self):
        self.stop.set()
        self.destroy()


# --------------------------------------------------------------------------
# 설정 창 (설비 목록 / 표시 · 연결)
# --------------------------------------------------------------------------
def _monitor_work_rect(cx, cy):
    """(cx, cy) 가 있는 모니터의 작업 영역 (left, top, right, bottom). 윈도우 전용, 실패하면 None."""
    if sys.platform != "win32":
        return None
    try:
        class MI(ctypes.Structure):
            _fields_ = [("cbSize", wintypes.DWORD), ("rcMonitor", wintypes.RECT),
                        ("rcWork", wintypes.RECT), ("dwFlags", wintypes.DWORD)]
        u = ctypes.windll.user32
        u.MonitorFromPoint.argtypes = [wintypes.POINT, wintypes.DWORD]
        u.MonitorFromPoint.restype = wintypes.HANDLE
        u.GetMonitorInfoW.argtypes = [wintypes.HANDLE, ctypes.POINTER(MI)]
        hm = u.MonitorFromPoint(wintypes.POINT(int(cx), int(cy)), 2)       # 2 = 가장 가까운 모니터
        mi = MI()
        mi.cbSize = ctypes.sizeof(MI)
        if hm and u.GetMonitorInfoW(hm, ctypes.byref(mi)):
            r = mi.rcWork
            return r.left, r.top, r.right, r.bottom
    except Exception:
        pass
    return None


def round_window(win, radius, full=False):
    """제목표시줄 없는 창의 모서리를 둥글게 (Windows).
    Windows 11: DWM 둥근 모서리(부드러움, 반지름은 윈도우가 정함). Windows 10: 창 영역 자르기(계단짐)."""
    if sys.platform != "win32":
        return
    try:
        u = ctypes.WinDLL("user32", use_last_error=True)
        g = ctypes.WinDLL("gdi32", use_last_error=True)
        u.GetParent.argtypes, u.GetParent.restype = [wintypes.HWND], wintypes.HWND
        u.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
        u.SetWindowRgn.argtypes = [wintypes.HWND, ctypes.c_void_p, wintypes.BOOL]
        g.CreateRoundRectRgn.argtypes = [ctypes.c_int] * 6
        g.CreateRoundRectRgn.restype = ctypes.c_void_p
        hwnd = u.GetParent(win.winfo_id()) or win.winfo_id()
        ok11 = False
        if sys.getwindowsversion().build >= 22000:
            dwm = ctypes.WinDLL("dwmapi")
            dwm.DwmSetWindowAttribute.argtypes = [wintypes.HWND, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD]
            pref = ctypes.c_int(1 if full else 2)           # 1 = 둥글게 안 함, 2 = 둥글게
            ok11 = dwm.DwmSetWindowAttribute(hwnd, 33, ctypes.byref(pref), 4) == 0
        if ok11 or full:
            u.SetWindowRgn(hwnd, None, True)
            return
        r = wintypes.RECT()
        if not u.GetWindowRect(hwnd, ctypes.byref(r)):
            return
        d = 2 * radius
        u.SetWindowRgn(hwnd, g.CreateRoundRectRgn(0, 0, r.right - r.left + 1, r.bottom - r.top + 1, d, d), True)
    except Exception:
        pass


def center_on(master, w, h):
    """w x h 창을 '파일 전송(메인) 창이 있는 모니터'의 가운데에 놓는 좌상단 좌표.
    메인 창이 안 보이면 모니터 창, 그것도 아니면 기본 화면."""
    ref = None
    for cand in (getattr(master, "tw", None), master):
        try:
            if cand is not None and cand.winfo_viewable():
                ref = cand
                break
        except tk.TclError:
            pass
    if ref is not None:
        cx = ref.winfo_rootx() + ref.winfo_width() // 2
        cy = ref.winfo_rooty() + ref.winfo_height() // 2
    else:
        cx, cy = master.winfo_screenwidth() // 2, master.winfo_screenheight() // 2
    rect = _monitor_work_rect(cx, cy) or (0, 0, master.winfo_screenwidth(), master.winfo_screenheight())
    l, t, r, b = rect
    mcx, mcy = (l + r) // 2, (t + b) // 2
    return max(l, min(mcx - w // 2, r - w)), max(t, min(mcy - h // 2, b - h))


class SiteManager(tk.Toplevel):
    """설비 추가/수정/삭제/순서 변경, 장비 찾기, 전송 설정. '적용'을 눌러야 저장되고 반영된다."""

    def __init__(self, master, on_apply=None):
        super().__init__(master)
        self.on_apply = on_apply
        self.monitor = master
        self.title("FTP Lite 설정")
        self.configure(bg=C_CARD, highlightthickness=1, highlightbackground="#9a9aa2", highlightcolor="#9a9aa2")
        h = min(640, self.winfo_screenheight() - 80)
        x, y = center_on(master, 800, h)
        self.overrideredirect(True)                  # 제목표시줄 없음
        self.resizable(False, False)                 # 크기 고정
        self.geometry(f"800x{h}+{x}+{y}")

        cfg = load_app_cfg()
        self.entries = []                 # [{"uid","name","host","user","password"}] 표시 순서대로
        self._next_uid = 1
        for name, s_ in load_sites().items():
            self._new_entry(name, s_["host"], s_["user"], s_["password"])
        self.cur = None                   # 선택된 설비의 uid
        self._loading = False
        self._applied = False
        self.name_var = tk.StringVar()
        self.host_var = tk.StringVar()
        self.user_var = tk.StringVar()
        self.pw_var = tk.StringVar()
        self._capturing = False           # 단축키 입력 중
        self.reconnect_var = tk.StringVar(value=str(cfg_seconds(cfg, "reconnect_sec", DEFAULT_RECONNECT_SEC)))
        self.keepalive_var = tk.StringVar(value=str(cfg_seconds(cfg, "keepalive_sec", DEFAULT_KEEPALIVE_SEC)))
        self.tw_hotkey_var = tk.StringVar(value=tw_hotkey_cfg(cfg))
        self.ncdata_var = tk.StringVar(value=str(pc_default_dir(cfg)))
        self.editor_var = tk.StringVar(value=str(cfg.get("editor_path", "") or ""))
        self.automin_var = tk.BooleanVar(value=bool(cfg.get("auto_min_done", True)))
        self.refresh_shown_var = tk.StringVar(value=str(cfg_seconds(cfg, "refresh_shown_sec", 1)))
        self.refresh_hidden_var = tk.StringVar(value=str(cfg_seconds(cfg, "refresh_hidden_sec", 30)))
        self.autostart_var = tk.BooleanVar(value=autostart_get())
        for v in (self.name_var, self.host_var, self.user_var, self.pw_var):
            v.trace_add("write", self._on_edit)
        self._test_result = None
        self._orig = self._snapshot()
        self.finder = None

        self.build_ui()
        self.refresh_list(select=self.entries[0]["uid"] if self.entries else None)
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.bind("<Escape>", self._esc)
        self.bind("<KeyPress>", self._on_key)
        self.after(50, self._grab_focus)
        self.after(80, lambda: round_window(self, 8))

    def _esc(self, _e=None):
        if self._capturing:
            self._end_capture()
        else:
            self.close()

    # ---- 데이터 ----
    def _new_entry(self, name, host="", user="", password="", index=None):
        e = {"uid": self._next_uid, "name": name, "host": host, "user": user, "password": password}
        self._next_uid += 1
        if index is None:
            self.entries.append(e)
        else:
            self.entries.insert(index, e)
        return e

    def _snapshot(self):
        return json.dumps([[[e["name"], e["host"], e["user"], e["password"]] for e in self.entries],
                           self.reconnect_var.get(), self.keepalive_var.get(), self.tw_hotkey_var.get(),
                           self.ncdata_var.get(), self.editor_var.get(), self.automin_var.get(),
                           self.refresh_shown_var.get(), self.refresh_hidden_var.get(), self.autostart_var.get()],
                          ensure_ascii=False)

    # ---- UI ----
    def build_ui(self):
        top = tk.Frame(self, bg=C_CARD)
        top.pack(fill="x", padx=22, pady=(18, 10))
        tk.Label(top, text="설정", font=ui_font(20, True), bg=C_CARD, fg=C_TEXT).pack(side="left")
        self.tabs = Segmented(top, ["설비", "설정"], self._switch_tab, width=200, outer=C_CARD)
        self.tabs.pack(side="right")

        foot = tk.Frame(self, bg=C_CARD)
        foot.pack(side="bottom", fill="x", padx=22, pady=(10, 18))
        RoundButton(foot, "닫기", self.close, width=80, height=34, outer=C_CARD).pack(side="right")
        RoundButton(foot, "적용", self.apply, kind="primary", width=92, height=34, outer=C_CARD).pack(side="right", padx=8)
        tk.Label(foot, text=f"FTP Lite  v{APP_VERSION}\n설정 파일  {DATA_PATH}", font=ui_font(8), bg=C_CARD,
                 fg=C_HINT, wraplength=500, justify="left", anchor="w").pack(side="left")

        body = tk.Frame(self, bg=C_CARD)
        body.pack(fill="both", expand=True, padx=22)
        self.page_sites = tk.Frame(body, bg=C_CARD)
        self.page_misc = ScrollPage(body)
        self._build_sites_page(self.page_sites)
        self._build_misc_page(self.page_misc.inner)
        self._switch_tab(0)

    def _switch_tab(self, idx):
        self.tabs.select(idx)
        pages = (self.page_sites, self.page_misc)
        for p in pages:
            p.pack_forget()
        pages[idx].pack(fill="both", expand=True)

    def _build_sites_page(self, page):
        page.columnconfigure(1, weight=1)
        page.rowconfigure(0, weight=1)

        wrap, left = fsection(page, "설비 목록", expand=True, width=270)
        wrap.configure(width=270)
        wrap.pack_propagate(False)                  # 목록이 아무리 길어도 왼쪽 폭 고정
        wrap.grid(row=0, column=0, sticky="nsew", padx=(0, 18))
        page.columnconfigure(0, minsize=270, weight=0)
        self.tree = ttk.Treeview(left, show="tree", selectmode="browse", style="Apple.Treeview")
        self.tree.column("#0", width=200, stretch=True)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        grid = tk.Frame(left, bg=C_CARD)
        grid.pack(pady=(12, 0))
        for i, (text, cmd, kind) in enumerate((
                ("＋ 장비 추가", self.new_site, "primary"), ("장비 찾기", self.open_finder, "secondary"),
                ("복제", self.duplicate_site, "secondary"), ("삭제", self.delete_site, "danger"),
                ("▲ 위로", lambda: self.move(-1), "secondary"), ("▼ 아래로", lambda: self.move(1), "secondary"))):
            RoundButton(grid, text, cmd, kind=kind, width=122, height=34).grid(
                row=i // 2, column=i % 2, padx=(0, 8) if i % 2 == 0 else 0, pady=(0, 8) if i < 4 else 0)

        wrap, form = fsection(page, "선택한 설비")
        wrap.grid(row=0, column=1, sticky="nsew")
        form.columnconfigure(1, weight=1)
        self.entries_w = []
        for r, (label, var, show) in enumerate((("호스트(IP)", self.host_var, None),
                                                ("장비명", self.name_var, None),
                                                ("사용자", self.user_var, None),
                                                ("비밀번호", self.pw_var, None))):
            tk.Label(form, text=label, font=ui_font(10), bg=C_CARD, fg=C_SUB, anchor="center", width=8).grid(
                row=r, column=0, sticky="ew", padx=(0, 8), pady=6)
            ent = field_entry(form, var, show=show, justify="center")
            ent.grid(row=r, column=1, sticky="ew", pady=6, ipady=6)
            self.entries_w.append(ent)
        self.host_entry, self.name_entry, self.user_entry, self.pw_entry = self.entries_w

        line = tk.Frame(form, bg=C_LINE, height=1)
        line.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(10, 12))
        act = tk.Frame(form, bg=C_CARD)
        act.grid(row=6, column=0, columnspan=2, sticky="ew")
        self.test_btn = RoundButton(act, "연결 확인", self.test_connection, width=100)
        self.test_btn.pack(side="left")
        tk.Label(act, text=f"FTP {FTP_PORT}", font=ui_font(8), bg=C_CARD, fg=C_HINT).pack(side="left", padx=10)
        self.result = tk.Text(form, height=5, width=10, wrap="word", relief="flat", bd=0, bg="#f7f7fa", fg=C_SUB,
                              font=ui_font(9), padx=12, pady=10, highlightthickness=0, state="disabled",
                              cursor="arrow")
        self.result.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        self.result.tag_configure("ok", foreground="#1c8a3a")
        self.result.tag_configure("bad", foreground=C_RED)
        self.result.tag_configure("hint", foreground=C_SUB)



    def _build_misc_page(self, page):
        wrap, ftp = fsection(page, "파일 전송")
        wrap.pack(fill="x")
        right = self._row(ftp, "내 PC 기본 폴더")
        RoundButton(right, "찾기", self._pick_pc, width=70).pack(side="right")
        field_entry(right, self.ncdata_var, width=22).pack(side="right", padx=(0, 8), ipady=4)
        right = self._row(ftp, "전송이 모두 성공하면 창 자동 최소화", last=True)
        Switch(right, self.automin_var).pack()

        wrap, conn = fsection(page, "연결  (모든 설비 공통)")
        wrap.pack(fill="x", pady=(10, 0))
        right = self._row(conn, "다시 연결 간격")
        tk.Label(right, text="초", font=ui_font(9), bg=C_CARD, fg=C_SUB).pack(side="right", padx=(6, 0))
        field_entry(right, self.reconnect_var, width=6).pack(side="right", ipady=4)
        right = self._row(conn, "살아있는지 확인 간격", last=True)
        tk.Label(right, text="초", font=ui_font(9), bg=C_CARD, fg=C_SUB).pack(side="right", padx=(6, 0))
        field_entry(right, self.keepalive_var, width=6).pack(side="right", ipady=4)

        wrap, rf = fsection(page, "목록 새로고침 주기")
        wrap.pack(fill="x", pady=(10, 0))
        for label, var, last in (("창이 보일 때", self.refresh_shown_var, False), ("창이 숨겨졌을 때", self.refresh_hidden_var, True)):
            right = self._row(rf, label, last=last)
            tk.Label(right, text="초", font=ui_font(9), bg=C_CARD, fg=C_SUB).pack(side="right", padx=(6, 0))
            field_entry(right, var, width=6).pack(side="right", ipady=4)

        wrap, ed = fsection(page, "파일 열기 프로그램  (오른쪽 클릭)")
        wrap.pack(fill="x", pady=(10, 0))
        right = self._row(ed, "프로그램 경로", "비워 두면 메모장", last=True)
        RoundButton(right, "기본", lambda: self.editor_var.set(""), width=60).pack(side="right")
        RoundButton(right, "찾기", self._pick_editor, width=60).pack(side="right", padx=(0, 6))
        field_entry(right, self.editor_var, width=18).pack(side="right", padx=(0, 6), ipady=4)

        wrap, key = fsection(page, "단축키")
        wrap.pack(fill="x", pady=(10, 0))
        right = self._row(key, "보이기 / 숨기기", last=True)
        self.tw_hk_btn = RoundButton(right, self.tw_hotkey_var.get(), self._start_capture, width=150)
        self.tw_hk_btn.pack()

        if sys.platform == "win32":
            wrap, boot = fsection(page, "시작")
            wrap.pack(fill="x", pady=(10, 0))
            right = self._row(boot, "윈도우 시작 시 자동 실행", last=True)
            Switch(right, self.autostart_var).pack()

        wrap, info = fsection(page, "정보")
        wrap.pack(fill="x", pady=(10, 0))
        right = self._row(info, "제작자", last=True)
        tk.Label(right, text="정성철 · sc9727@naver.com", font=ui_font(9), bg=C_CARD, fg=C_SUB).pack()

    # ---- 단축키 입력 ----
    def _start_capture(self):
        self._capturing = True
        self.tw_hk_btn.config(text="키를 누르세요…")
        self.focus_set()

    def _end_capture(self, text=None):
        if not self._capturing:
            return
        self._capturing = False
        if text:
            self.tw_hotkey_var.set(text)
        self.tw_hk_btn.config(text=self.tw_hotkey_var.get())

    def _on_key(self, e):
        if not self._capturing:
            return
        ks = e.keysym
        if ks.startswith(("Control", "Shift", "Alt", "Meta", "Super", "Win", "Caps", "Num", "ISO", "Hyper")):
            return                                              # 수정키만 누른 상태는 기다린다
        alt = 0x20000 if sys.platform == "win32" else 0x8
        mods = (["Ctrl"] if e.state & 4 else []) + (["Alt"] if e.state & alt else []) + \
               (["Shift"] if e.state & 1 else [])
        if len(ks) == 1 and ks.isalnum() and ks.isascii():
            key = ks.upper()
        elif re.fullmatch(r"F\d{1,2}", ks):
            key = ks
        else:
            messagebox.showinfo("단축키", "영문자, 숫자, F1~F12 키만 쓸 수 있습니다.", parent=self)
            self._end_capture()
            return
        text = "+".join(mods + [key])
        if parse_hotkey(text) is None:
            messagebox.showinfo("단축키", "Ctrl 또는 Alt 키와 함께 눌러야 합니다.\n예) Ctrl+Alt+F", parent=self)
            self._end_capture()
            return
        self._end_capture(text)

    # ---- 목록 ----
    def refresh_list(self, select=None):
        self._loading = True
        try:
            self.tree.delete(*self.tree.get_children())
            for e in self.entries:
                self.tree.insert("", "end", iid=str(e["uid"]), text="  " + (e["name"] or "(이름 없음)"))
            if select is not None and self._entry(select):
                self.tree.selection_set(str(select))
                self.tree.see(str(select))
        finally:
            self._loading = False
        self.load_form(select if select is not None and self._entry(select) else None)

    def new_site(self, host=""):
        if not self._check_current():
            return None
        e = self._new_entry("", host)                     # 장비명은 비워 두고 직접 입력
        self.refresh_list(select=e["uid"])
        self._focus_name()
        return e

    def duplicate_site(self):
        if self.cur is None or not self._check_current():
            return
        src = self._entry(self.cur)
        e = self._new_entry(self._unique(f"{src['name']} 복사본"), src["host"], src["user"], src["password"],
                            index=self._index(self.cur) + 1)
        self.refresh_list(select=e["uid"])
        self._focus_name()

    # ---- 연결 확인 (FTP 만) ----
    def test_connection(self):
        host = self.host_var.get().strip()
        if self.cur is None or not host:
            messagebox.showwarning("연결 확인", "호스트(IP)를 입력하세요.", parent=self)
            return
        self.test_btn.config(state="disabled")
        self._set_result(f"{host} 확인 중...")
        self._test_result = None
        threading.Thread(target=self._test_worker, args=(host, self.user_var.get(), self.pw_var.get()),
                         daemon=True).start()
        self._test_poll()

    def _test_worker(self, host, user, password):
        lines, ok = [], True
        t0 = time.perf_counter()
        try:
            ftp = FTP()
            ftp.connect(host, FTP_PORT, timeout=5)
            try:
                ftp.login(user or "anonymous", password)
                lines.append(f"✔ FTP({FTP_PORT}) 로그인 성공  ({(time.perf_counter() - t0) * 1000:.0f} ms)")
            except error_perm as e:
                ok = False
                lines.append(f"✘ FTP({FTP_PORT}) 연결은 되지만 로그인 실패: {e}\n   → 사용자/비밀번호를 확인하세요")
            finally:
                try:
                    ftp.quit()
                except all_errors:
                    ftp.close()
        except (all_errors + (EOFError,)) as e:
            ok = False
            lines.append(f"✘ FTP({FTP_PORT}) 연결 실패: {e}\n   → IP·랜선·장비 전원을 확인하세요")
        self._test_result = (ok, "\n".join(lines))

    def _check_all(self):
        """모든 설비의 장비명이 채워져 있고 겹치지 않는지 확인 (문제 있으면 그 설비를 선택해 안내)."""
        if not self._check_current():
            return False
        seen = set()
        for e in self.entries:
            if not e["name"] or e["name"] in seen:
                self.refresh_list(select=e["uid"])
                self._check_current()
                return False
            seen.add(e["name"])
        return True

    # ---- 적용 / 닫기 ----
    def _collect(self):
        """화면에 입력된 내용 -> (설비 목록, 설정)."""
        cfg = load_app_cfg()
        cfg["reconnect_sec"] = cfg_seconds({"v": self.reconnect_var.get()}, "v", DEFAULT_RECONNECT_SEC)
        cfg["keepalive_sec"] = cfg_seconds({"v": self.keepalive_var.get()}, "v", DEFAULT_KEEPALIVE_SEC)
        cfg["tw_hotkey"] = self.tw_hotkey_var.get()
        cfg["pc_dir"] = self.ncdata_var.get().strip()
        cfg["editor_path"] = self.editor_var.get().strip()
        cfg["auto_min_done"] = bool(self.automin_var.get())
        cfg["refresh_shown_sec"] = cfg_seconds({"v": self.refresh_shown_var.get()}, "v", 1)
        cfg["refresh_hidden_sec"] = cfg_seconds({"v": self.refresh_hidden_var.get()}, "v", 30)
        sites = {e["name"]: {"host": e["host"], "user": e["user"], "password": e["password"], "ftp": True}
                 for e in self.entries}
        return sites, cfg

    def apply(self):
        if not self._check_all():
            return
        sites, cfg = self._collect()
        try:
            save_sites(sites)
        except OSError as e:
            messagebox.showerror("저장 실패", f"설비 목록을 저장하지 못했습니다.\n{e}", parent=self)
            return
        save_app_cfg(cfg)
        if self.autostart_var.get() != autostart_get():
            autostart_set(self.autostart_var.get())
        self._applied = True
        self._close_now()
        if self.on_apply:
            self.on_apply()

    def close(self):
        if self._snapshot() != self._orig and not messagebox.askyesno(
                "닫기", "바꾼 내용을 적용하지 않고 닫을까요?", parent=self):
            return
        self._close_now()

    def _close_now(self):
        if self.finder is not None and self.finder.winfo_exists():
            self.finder.close()
        self.destroy()

    def _grab_focus(self):
        try:
            self.lift()
            self.focus_force()
        except tk.TclError:
            pass

    def _entry(self, uid):
        return next((e for e in self.entries if e["uid"] == uid), None)

    def _index(self, uid):
        return next((i for i, e in enumerate(self.entries) if e["uid"] == uid), -1)

    def _unique(self, base):
        names = {e["name"] for e in self.entries}
        name, n = base, 2
        while name in names:
            name, n = f"{base} ({n})", n + 1
        return name

    # ---- UI ----

    def _row(self, parent, label, sub=None, last=False):
        f = tk.Frame(parent, bg=C_CARD)
        f.pack(fill="x", pady=5)
        lab = tk.Frame(f, bg=C_CARD)
        lab.pack(side="left")
        tk.Label(lab, text=label, font=ui_font(10), bg=C_CARD, fg=C_TEXT, anchor="w").pack(anchor="w")
        if sub:
            tk.Label(lab, text=sub, font=ui_font(8), bg=C_CARD, fg=C_HINT, anchor="w", justify="left",
                     wraplength=330).pack(anchor="w")
        right = tk.Frame(f, bg=C_CARD)
        right.pack(side="right")
        if not last:
            tk.Frame(parent, height=1, bg=C_LINE).pack(fill="x")
        return right

    def _pick_pc(self):
        d = filedialog.askdirectory(parent=self, title="내 PC 기본 폴더 선택", initialdir=self.ncdata_var.get() or None)
        if d:
            self.ncdata_var.set(os.path.normpath(d))

    def _pick_editor(self):
        kw = {"filetypes": [("프로그램", "*.exe"), ("모든 파일", "*.*")]} if IS_WINDOWS else {}
        f = filedialog.askopenfilename(parent=self, title="파일을 열 프로그램 선택", **kw)
        if f:
            self.editor_var.set(os.path.normpath(f))

    def _set_form_state(self):
        state = "normal" if self.cur is not None else "disabled"
        for e in self.entries_w:
            e.config(state=state)

    # ---- 목록 ----

    def load_form(self, uid):
        self._loading = True
        try:
            self.cur = uid
            e = self._entry(uid) if uid is not None else None
            self.name_var.set(e["name"] if e else "")
            self.host_var.set(e["host"] if e else "")
            self.user_var.set(e["user"] if e else "")
            self.pw_var.set(e["password"] if e else "")
        finally:
            self._loading = False
        self._set_form_state()
        self._set_result("")

    def _on_edit(self, *_):
        """입력칸을 고치면 바로 (메모리의) 선택된 설비에 반영. 파일 저장은 '적용'에서."""
        if self._loading or self.cur is None:
            return
        e = self._entry(self.cur)
        if not e:
            return
        e["name"] = self.name_var.get().strip()
        e["host"] = self.host_var.get().strip()
        e["user"] = self.user_var.get()
        e["password"] = self.pw_var.get()
        try:
            self.tree.item(str(self.cur), text="  " + (e["name"] or "(이름 없음)"))
        except tk.TclError:
            pass

    def _check_current(self):
        """선택된 설비의 이름이 비어 있거나 겹치면 안내하고 False."""
        e = self._entry(self.cur) if self.cur is not None else None
        if not e:
            return True
        if not e["name"]:
            messagebox.showwarning("장비명", "장비명을 입력하세요.", parent=self)
        elif any(o["name"] == e["name"] and o["uid"] != e["uid"] for o in self.entries):
            messagebox.showwarning("장비명", f"'{e['name']}' 이름의 설비가 이미 있습니다.\n다른 이름을 입력하세요.", parent=self)
        else:
            return True
        self.name_entry.focus_set()
        return False

    def on_select(self, _e=None):
        if self._loading:
            return
        sel = self.tree.selection()
        uid = int(sel[0]) if sel else None
        if uid == self.cur:
            return
        if not self._check_current():
            self._loading = True                # 이름을 고칠 때까지 원래 선택을 유지
            try:
                self.tree.selection_set(str(self.cur))
            finally:
                self._loading = False
            return
        self.load_form(uid)

    def move(self, step):
        if self.cur is None or not self._check_current():
            return
        i = self._index(self.cur)
        j = i + step
        if not 0 <= j < len(self.entries):
            return
        self.entries[i], self.entries[j] = self.entries[j], self.entries[i]
        self.refresh_list(select=self.cur)

    def _focus_name(self):
        def go():
            self.name_entry.focus_set()
            self.name_entry.select_range(0, "end")
        self.after(80, go)

    def delete_site(self):
        if self.cur is None:
            return
        e = self._entry(self.cur)
        if not messagebox.askyesno("삭제", f"'{e['name'] or '(이름 없음)'}' 설비를 삭제할까요?", parent=self):
            return
        i = self._index(self.cur)
        self.entries.pop(i)
        self.cur = None
        self.refresh_list(select=self.entries[min(i, len(self.entries) - 1)]["uid"] if self.entries else None)

    # ---- 연결 확인 ----

    def _set_result(self, text):
        self.result.config(state="normal")
        self.result.delete("1.0", "end")
        for line in text.splitlines(True):
            tag = "ok" if line.startswith("✔") else "bad" if line.startswith("✘") else "hint"
            self.result.insert("end", line, tag)
        self.result.config(state="disabled")

    def _test_poll(self):
        if not self.winfo_exists():
            return
        if self._test_result is None:
            self.after(100, self._test_poll)
            return
        _ok, text = self._test_result
        self.test_btn.config(state="normal")
        self._set_result(text)

    # ---- 장비 찾기 ----

    def open_finder(self):
        if self.finder is not None and self.finder.winfo_exists():
            self.finder.lift()
            return
        hosts = [e["host"] for e in self.entries if e["host"]]
        self.finder = DeviceFinder(self, local_prefixes(hosts), self._name_of_host, self.add_found)

    def _name_of_host(self, ip):
        return next((e["name"] for e in self.entries if e["host"] == ip), None)

    def add_found(self, ips):
        """장비 찾기에서 체크한 IP 들을 새 설비로 추가 (장비명은 비워 둠, 이미 있는 IP 는 건너뜀)."""
        if isinstance(ips, str):
            ips = [ips]
        if not self._check_current():
            return False
        first = None
        for ip in ips:
            e = next((x for x in self.entries if x["host"] == ip), None) or self._new_entry("", ip)
            first = first or e
        if first:
            self.refresh_list(select=first["uid"])
            self._focus_name()
        return True


# --------------------------------------------------------------------------
# Windows 시스템 트레이 아이콘 + 전역 단축키 (ctypes, 추가 설치 없음)
# --------------------------------------------------------------------------
IS_WINDOWS = sys.platform == "win32"

if IS_WINDOWS:
    import ctypes
    from ctypes import wintypes

    _user32 = ctypes.WinDLL("user32", use_last_error=True)
    _shell32 = ctypes.WinDLL("shell32", use_last_error=True)
    _gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
    _kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    _LRESULT = ctypes.c_ssize_t
    _WNDPROC = ctypes.WINFUNCTYPE(_LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

    class _WNDCLASSW(ctypes.Structure):
        _fields_ = [("style", wintypes.UINT), ("lpfnWndProc", _WNDPROC),
                    ("cbClsExtra", ctypes.c_int), ("cbWndExtra", ctypes.c_int),
                    ("hInstance", wintypes.HINSTANCE), ("hIcon", wintypes.HICON),
                    ("hCursor", wintypes.HANDLE), ("hbrBackground", wintypes.HBRUSH),
                    ("lpszMenuName", wintypes.LPCWSTR), ("lpszClassName", wintypes.LPCWSTR)]

    class _GUID(ctypes.Structure):
        _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD),
                    ("Data3", wintypes.WORD), ("Data4", ctypes.c_ubyte * 8)]

    class _NOTIFYICONDATAW(ctypes.Structure):
        _fields_ = [("cbSize", wintypes.DWORD), ("hWnd", wintypes.HWND), ("uID", wintypes.UINT),
                    ("uFlags", wintypes.UINT), ("uCallbackMessage", wintypes.UINT),
                    ("hIcon", wintypes.HICON), ("szTip", wintypes.WCHAR * 128),
                    ("dwState", wintypes.DWORD), ("dwStateMask", wintypes.DWORD),
                    ("szInfo", wintypes.WCHAR * 256), ("uVersion", wintypes.UINT),
                    ("szInfoTitle", wintypes.WCHAR * 64), ("dwInfoFlags", wintypes.DWORD),
                    ("guidItem", _GUID), ("hBalloonIcon", wintypes.HICON)]

    class _BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG), ("biHeight", wintypes.LONG),
                    ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD),
                    ("biCompression", wintypes.DWORD), ("biSizeImage", wintypes.DWORD),
                    ("biXPelsPerMeter", wintypes.LONG), ("biYPelsPerMeter", wintypes.LONG),
                    ("biClrUsed", wintypes.DWORD), ("biClrImportant", wintypes.DWORD)]

    class _ICONINFO(ctypes.Structure):
        _fields_ = [("fIcon", wintypes.BOOL), ("xHotspot", wintypes.DWORD), ("yHotspot", wintypes.DWORD),
                    ("hbmMask", wintypes.HBITMAP), ("hbmColor", wintypes.HBITMAP)]

    class _MSG(ctypes.Structure):
        _fields_ = [("hwnd", wintypes.HWND), ("message", wintypes.UINT), ("wParam", wintypes.WPARAM),
                    ("lParam", wintypes.LPARAM), ("time", wintypes.DWORD), ("pt", wintypes.POINT)]

    def _sig(fn, res, *args):
        fn.restype = res
        fn.argtypes = list(args)

    _sig(_user32.DefWindowProcW, _LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
    _sig(_user32.RegisterClassW, wintypes.ATOM, ctypes.POINTER(_WNDCLASSW))
    _sig(_user32.CreateWindowExW, wintypes.HWND, wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR,
         wintypes.DWORD, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
         wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, wintypes.LPVOID)
    _sig(_user32.DestroyWindow, wintypes.BOOL, wintypes.HWND)
    _sig(_user32.GetMessageW, wintypes.BOOL, ctypes.POINTER(_MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT)
    _sig(_user32.TranslateMessage, wintypes.BOOL, ctypes.POINTER(_MSG))
    _sig(_user32.DispatchMessageW, _LRESULT, ctypes.POINTER(_MSG))
    _sig(_user32.PostMessageW, wintypes.BOOL, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
    _sig(_user32.PostQuitMessage, None, ctypes.c_int)
    _sig(_user32.RegisterWindowMessageW, wintypes.UINT, wintypes.LPCWSTR)
    _sig(_user32.CreatePopupMenu, wintypes.HMENU)
    _sig(_user32.AppendMenuW, wintypes.BOOL, wintypes.HMENU, wintypes.UINT, ctypes.c_size_t, wintypes.LPCWSTR)
    _sig(_user32.TrackPopupMenu, wintypes.BOOL, wintypes.HMENU, wintypes.UINT, ctypes.c_int, ctypes.c_int,
         ctypes.c_int, wintypes.HWND, wintypes.LPVOID)
    _sig(_user32.DestroyMenu, wintypes.BOOL, wintypes.HMENU)
    _sig(_user32.SetForegroundWindow, wintypes.BOOL, wintypes.HWND)
    _sig(_user32.GetCursorPos, wintypes.BOOL, ctypes.POINTER(wintypes.POINT))
    _sig(_user32.RegisterHotKey, wintypes.BOOL, wintypes.HWND, ctypes.c_int, wintypes.UINT, wintypes.UINT)
    _sig(_user32.UnregisterHotKey, wintypes.BOOL, wintypes.HWND, ctypes.c_int)
    _sig(_user32.CreateIconIndirect, wintypes.HICON, ctypes.POINTER(_ICONINFO))
    _sig(_user32.DestroyIcon, wintypes.BOOL, wintypes.HICON)
    _sig(_user32.GetSystemMetrics, ctypes.c_int, ctypes.c_int)
    _sig(_shell32.Shell_NotifyIconW, wintypes.BOOL, wintypes.DWORD, ctypes.POINTER(_NOTIFYICONDATAW))
    _sig(_gdi32.CreateDIBSection, wintypes.HBITMAP, wintypes.HDC, ctypes.c_void_p, wintypes.UINT,
         ctypes.POINTER(ctypes.c_void_p), wintypes.HANDLE, wintypes.DWORD)
    _sig(_gdi32.CreateBitmap, wintypes.HBITMAP, ctypes.c_int, ctypes.c_int, wintypes.UINT, wintypes.UINT,
         ctypes.c_void_p)
    _sig(_gdi32.DeleteObject, wintypes.BOOL, wintypes.HGDIOBJ)
    _sig(_user32.SetLayeredWindowAttributes, wintypes.BOOL, wintypes.HWND, wintypes.DWORD, ctypes.c_ubyte,
         wintypes.DWORD)
    if hasattr(_user32, "GetWindowLongPtrW"):              # 64비트
        _GetWindowLong, _SetWindowLong = _user32.GetWindowLongPtrW, _user32.SetWindowLongPtrW
    else:                                                   # 32비트
        _GetWindowLong, _SetWindowLong = _user32.GetWindowLongW, _user32.SetWindowLongW
    _sig(_GetWindowLong, ctypes.c_ssize_t, wintypes.HWND, ctypes.c_int)
    _sig(_SetWindowLong, ctypes.c_ssize_t, wintypes.HWND, ctypes.c_int, ctypes.c_ssize_t)
    _sig(_kernel32.GetModuleHandleW, wintypes.HMODULE, wintypes.LPCWSTR)
    _sig(_user32.FindWindowW, wintypes.HWND, wintypes.LPCWSTR, wintypes.LPCWSTR)
    _sig(_kernel32.CreateMutexW, wintypes.HANDLE, wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR)

    _WM_DESTROY, _WM_CLOSE, _WM_NULL, _WM_HOTKEY = 0x0002, 0x0010, 0x0000, 0x0312
    _WM_LBUTTONUP, _WM_LBUTTONDBLCLK, _WM_RBUTTONUP = 0x0202, 0x0203, 0x0205
    _WM_TRAY = 0x8001            # WM_APP + 1
    _WM_SHOW = 0x8002            # WM_APP + 2
    _WM_SETKEY = 0x8003          # WM_APP + 3 (단축키 다시 등록)
    _TRAY_CLASS = "FtpLiteTray"
    _NIM_ADD, _NIM_MODIFY, _NIM_DELETE = 0, 1, 2
    _NIF_MESSAGE, _NIF_ICON, _NIF_TIP = 1, 2, 4
    _MF_STRING, _MF_CHECKED, _MF_SEPARATOR = 0x0, 0x8, 0x800
    _TPM_RIGHTBUTTON, _TPM_RETURNCMD, _TPM_BOTTOMALIGN = 0x0002, 0x0100, 0x0020
    _MOD_ALT, _MOD_CONTROL, _MOD_NOREPEAT = 0x0001, 0x0002, 0x4000
    _HOTKEY_ID = 1
_HOTKEY_ID2 = 2


def single_instance(name="FtpLite_SingleInstance"):
    """이미 실행 중이면 False. (Windows 전용, 다른 OS 는 항상 True)"""
    if not IS_WINDOWS:
        return True
    global _INSTANCE_MUTEX
    _INSTANCE_MUTEX = _kernel32.CreateMutexW(None, False, name)
    return ctypes.get_last_error() != 183          # ERROR_ALREADY_EXISTS


def show_running_instance():
    """이미 실행 중인 프로그램에 '창 보이기'를 요청한다."""
    if IS_WINDOWS:
        hwnd = _user32.FindWindowW(_TRAY_CLASS, None)
        if hwnd:
            _user32.PostMessageW(hwnd, _WM_SHOW, 0, 0)
            return True
    return False


DEFAULT_HOTKEY = "Ctrl+Alt+F"


def parse_hotkey(text):
    """'Ctrl+Alt+V' -> (수정키 비트, 가상키 코드). 형식이 틀리면 None. (수정키: Alt 1, Ctrl 2, Shift 4, Win 8)"""
    mods, key = 0, None
    for part in str(text).replace(" ", "").split("+"):
        p = part.lower()
        if p in ("ctrl", "control"):
            mods |= 2
        elif p == "alt":
            mods |= 1
        elif p == "shift":
            mods |= 4
        elif p in ("win", "windows"):
            mods |= 8
        elif len(part) == 1 and part.isalnum() and part.isascii():
            key = ord(part.upper())
        elif re.fullmatch(r"[Ff](\d{1,2})", part) and 1 <= int(part[1:]) <= 12:
            key = 0x70 + int(part[1:]) - 1
        else:
            return None
    if key is None or not mods & (1 | 2 | 8):        # Shift 만으로는 글자 입력과 겹치므로 Ctrl/Alt/Win 필요
        return None
    return mods, key


def hotkey_tk_sequence(text):
    """트레이 아이콘이 없을 때 쓰는 tkinter 키 표기."""
    p = parse_hotkey(text)
    if not p:
        return None
    mods, key = p
    parts = (["Control"] if mods & 2 else []) + (["Alt"] if mods & 1 else []) + (["Shift"] if mods & 4 else [])
    name = f"F{key - 0x6f}" if 0x70 <= key <= 0x7B else chr(key).lower()
    return "<" + "-".join(parts + [name]) + ">"


def on_screen(x, y):
    """(x, y) 가 현재 연결된 모니터 영역(가상 화면) 안에 있는지."""
    if not IS_WINDOWS:
        return True
    vx, vy = _user32.GetSystemMetrics(76), _user32.GetSystemMetrics(77)
    vw, vh = _user32.GetSystemMetrics(78), _user32.GetSystemMetrics(79)
    return vx <= x < vx + vw - 40 and vy <= y < vy + vh - 40


_ICON_DATA = {
    16: "eNqlU0tLW1EQniT6C1wI+geEbn1BF0I1BqXR6yPdtETc+lip6yqK2iraXUW3ilpqm1xaogsDEaka3w9ocWHUasFKUVMJhDZ+zpx7o4m1brL4uN88vplz58whIgt5KzvJV/WL9MoY8ysG/oMrlSO5ojG0HeR3gXQND+iSIbmiEa2nIkKftBh5HuybDMlVGtaaPotXg02vUrDqt1wgsbjPyvjnnzhOnqegKTvofTHog4O/dtMuYbvMgOIOVS+uVdzjRIbvBUq/vMSzlT7kBVrhWGiHnW3nYhdy/E145G+GFuxGbqAF9LH8pn+anIfrupZfI/I3iuGQDw1bQ3gb+oxo7A82LvbgWuxE8GwXMyfrWD/fQ+PmkNKINq6vCfbiIPIT7pV+ZM/Ug8YLVa5zqQs0lodvv49Uj8njeYwdBUDvipCuV9/otaUenEbDGAlN8/lbYGPf9sUBnq8OwMq50n8nfIjRw1lkTdfBwvOSfzf0T+BeHcTX8HfVV82S57V7+QPutTegicfYj5ygNviK4wVqXmTO0GLOPpNrFsy1cT3jriQnn23xy7wKTW7zVhjx++5P7i3xbsUWv8SnEniS9nZ/7u6G1dwd4bYEnryLKe9vau8nxfd7DeR5JO4=",
    32: "eNrtl9tL1EEUx8fVwuqlevKh/oBeEqLAXgq8pIu63jKhHoKMCAnqUTHNLIPoanaVenBxFW/bXnRTM9FU1iCELl4yWy27IFQgWFGr+23O+e1Prb24gbo9OHD47Zw5Zz7nnJnf7G8ERIhQW2NaprCltQtr6qQwp7ikYJHExXPS3MRQG7FxSiOeHlklLCkVojUDojkdwpKKRWQrQnPS3MQgFjGJTc2k04vuLMjntJSZRWerQnMTQ2HpmW1OTueYTCnOJeN6xCFZxCS2Sdcl6+Ja0ry91UFhdslYpuR6YNnYs/uB6zC17NwVWVA0ljSEWdI9JMSc6lWvii9f0gXOl+eEMR6iLtpTTEkQ9THex+piFH9jgqde6gKJgW3ua5HYW4yioWqcHDSgwC35g5XY3JqN4y/ucr9wqIpt6En93AE9NDL2BPtp1pOO/HOe3ca2jhOcU8gCfK5hfTQq3rXDW4vuKYC/FmbJQPlYq9ex3P4KiIY9fuug8GNQ6rDC6ZpG/YceJHbnQ/ekBFp7MTbYDiC2pxDarjzUyTGyMX6yQ9udh6TeMwiROV55Y2G9YbwTO9uOcl2o/3PGiYjmg7K+ibyP/PGvORo55uKhGgh9JETtbkVMybKOWoiqKJwbrmObiyMmiOoonpdqp/qWDNdK361c948/vrIu0r0Ovmqg8m+MNrH92LcJtE30oePzS+QPVHJ+axv3IbQ+DpdHzGxT5mhCaEMc1jVmsW+Zm29434ntrYd5H7hkf9L5HRtt+3kPL5T/zVEbzzHDnkqrGn8sx2IRbt3Le/qSzJsa5Us+4dbMP3znt1+y9of6Snn9QwNYf7WGd0abESHn3dSSjfUUu1mn2PjhX3cotev5Moiifj2OPS/Hlkc5zPaV9998dQ+dH26AqNmlrDu9+6qN5F94bWSbq9J2Pl/1PftKrr9hB3NpzwTy/qvv3723D935t0AjWautGbOxK/xo3Bp9wDZkSz5rmD/nWz7WgjC5TyguTYDfU+r5o3WfP/H2Io/Y+bfU0RjZkC35cFxeff/1Wy517gylp9c182UTiG/g/z/+1syXTSC+K/IfSbC/P4P9/R3s+0ew71/Bvn8G+f79G4Rhjaw=",
    48: "eNrtmntMU1ccx08BH4wY96d/msXFv5c5s8SYbCgi8mgB0eArw2g0xsTE+K8ygizzD+czosm2ZImPaXGUlvoCH7AMfPH22UQpEuak6kYVkNp7f/v9fvfeUmtboG2kN/aQb057z++c+zmP3++cy60QagJh0D4KW8Hnwp6/W9SaWoTV5BY1RhkFH1gy35sYiIWYgrFSKhVJnB/7coqwG8vR/j9RVwjiXD5gGzAJ7Iro3sRALMREbMToz6z0xSDOFM0UNlODaFgBmGP9PC9KmjR2n4gBWYhJYWtgVmIG5k5GTUWbRnG1iOxHhGVS1kt4MROyKYyNzEzslKpzy0Uj9s2C5fHG/V4/kJFYiZlSVfZneH0Q/USKy3EPNg/ESszEbs2rYP/g9R7n7KM+4WVmYrcY28X5Ajk+fHUCPk3MxG4xvp7UGBlNbCV2vXEnlFBCH42SME4lWfNDq8Y0Pju13KDaU+5/XVMyKsVawJ9j0oc/slBLQ6t6mWq3dHyqzkZ7it05QcozQZzNAGFO5++R9sGg7HE8Rr89uQKXXZ1Q398B9S5UfzvrYn8b52UPT/MYWp5e5+91VI52VEdRh2rfyvnmjkoQpxdCzvVyLr+E7ZDNJdXG3PcXlD04BV9c3cZ90eY3Uv6+4RcQLl1xdfGcjzcd7b4I4sQ82Np5LKydjH+77p/APmROeB78+e+5e8ErS+CRvOAaGYCeoWfwZMgF3YNK/nNPHfN3uXv4uxOvkx3VIb2RPHz98eA/nO+8fxLEqa9hQ/thLh+R3sJr7zDceOmA2/86sK6b+SUso5TZ/D32YQnfIxL+h6/6fGOyvu0QpFmyYaZ9FaTVrmRNsy3nOqm1RTDDXgyp6A8lbQd9dRyv/oY0+0ouT6tdgTme3asWYVujNl1up+IX6BOzzhXDteddzC/JMhzvvcb25NPR8i+/tQfE7wuUe2k+Z8lVz+W5SmwxfwvFt/f66lB9aov9FW14HAP47+LcpeD16TZ8tjJ/AyWto2XkGxP15VD82+/8CnMurIW5l7fAnPrNMLtuI/uuFg+nWAvBULUY1rT89A5/MsdNI/sij2MAf+dAN66RxbzWaXz2PbKqa0jmuEB+HO36obY8sheGpREWrWvygU9oPahjr7GtDuBX+pjH7Qbyk586h/phWXMZ5KIqHGZs38PtU9pE8SoG64f4A9Ow18M+EA3/WxyTUOnPF/dgOs9bnsoUOb+M/BTXK3FujzovQGX3efjRcRam2grpmTqm/E/fvIT9j2pgBsUG9DXDBPeAUOvfdPMHECfnczwjBoFr3b9e5OvnGSxpLoWMplL4qmEHfIpxjNumebVGt3/5869r3Q8pyE5xgjhSAnwqEv53/TdDOUNgrEr2OytFw6/tX6RVLXvD+pLGRnZaHaofjL+k7YDPpnPAibwmZo6Ge6zzw3etB/De6WPwp7Odlqj++/zpGFeO+GwojvEeEaMzc7DzG2lR066wezlfx3Ky0+pQfYMfv2aThbFSs/kFzyCx5A95fqY9dMy5DTwbZ0VoE/vnl/Guy9Fnk3zf/hyJTUIJJZRQnL/P0Pf/z/X+/kLv74/0/v5O7+9P9f7+Wu+/H9D77zd0/vuZ/wEAsDO7",
    64: "eNrtm2lsVFUUx2+nRYl8cYkkBogxGlFAMFS/IIao2EKp86bTVhvUaJAo0YoIasUlNW4JifrBRCRxC0oiBbpMp3vtFiqVWBOiUEyjtCQ2Bm1rKa2UWd7xnPPmzbzOdJa2dyrte7f593X67nJ+dznnvpt5QoQlECmiuNjGf+fnpwq3kiWqHJ+JKuUXUekYFi4FLmuRjWQr2Uy2EwMlYiK2WEnnpuTOKRA1OZ0oEA25wFe3Ay57frLRaDMxEMtEjOPGPfD/QxuuxzLlojkfRJ0TRIWiCpfdp10VmBUy2kwMxEJMxGZkDe+Tb5UlotpxmvNXKF5RYffPGuaofYEMxEJMxEaMRuZiYWOVOa7DNfOraMylMp5Zzx3ZDx5mI0Zi1bkPBfxDhb1SG/c5yG7sA42xUlvrAXa3PS845+cqe8g3aGuBmCm1rEvD/vhZ1Ds1nzHX+YmRWImZ2MuVTI4Ts8m/y4gPxEzsLuULjpWmGHvDHGBmZHcpp0w7/sTuUi7Min1dMvaJxG42bkuWLFmyZMnSpJTiciSsqZSZqHwidczQnhn3jdmJi/LzXjN7akq43Ye4D9IqnWCrzEkaeyrWfU3NowlrHtpDZa+u2Typcpo2I4sjbrtU93x3Hj7HZYE48gCIso2cXyY792nZBljV8gIMei7AAGowhv6+dJ6v933/Bo9Lz+i5QLnhmOWMef64OACL6rcg03q4s3VH1Hb7Lw3D2X//gvaBLtjTXQormwu5jE3imtD4MyG9dSckknyqn68Zx4qZ/x/PCH9W8Sde0vOM+sZgScNTPKZ3te2CRJNX9cGbpw+wvbL6QOdf3foi2+fn32pcnk0dbzN/NE41Tj03Nm5l/vS2nQm1S+z+wL1nTuzleZAaWIOy+DXbQ6mkrx0+PVMN+3rrYF+Ppr09NXxd1lTI/B/95tLu9dZy3pK+oxG2H+k7hveqgvV8fMYN19Y+BqL0QeYPnz/vdR+GQmR8+eSX8NXZRp4vWh/4wa+q8OfYICxAv5AS8I1y+UN2LKY5evAetpP9j1G6D8dx4M+l6znvooYtEfP9lu+24b014+shu6O1W29ot2QtrzV9/PU5sOZoEd7PmLY/jMW/Gn3CfPTRV1U9gn44P6h57txgv1NcIl2B/0sr3wR3tGyP4L+77SW8l4V58oL5Y7W7En0x1XUl5mc+HOfukT6+5/F7OW/ej3u479OSyL+YfDSNA43x4ftDQttE2LzjeIZxZBn66HB+6kf2WQZbE+fX9kC9GAc0fh9fHcffTzo/re0iXIO7u76Boq798Mqp/fBq19dwU+PTHJONa08mP8cGnPc0v8XBtVDQ+QHf9av+YL7lOM+oPZt0/vgp+/g7vN5Tx/HI43+96wA82fkhbDvxCXzeW89z3jj2Pw39DjacHzJiYCx+ivXkd7x89cMY2kHXjI5i9k1y+WPvGzyqL/h5XftuKb4vnh1qmLyq5n0zO96aUX49nRsbAiete2xb1rNArHmY/cO7sLRhK9yOTEubng1qAcYD8sly/V8ojfguwnnvKAx5RnGv3A9t/SfhNfQ7N9Q+LpU9Hv/NFLd1f4/PHkGFscte//fi/F5YXQAL657A2Puw1ibtGdCO5Dz/TGwHxXJb+UZIczuZT1fKhPXI478V5xjH3Ao79zXVTXuGZJwFxLJjefPzCccYmfwrmrdDCj/rOpN+/hH5/BN6Epkqv15eryse/3TaTebzL43DZPnJ9vCU3rorKv9025V//hE6x7it6blJ81N8CD/vWNWyI6IeWe0m8/xrKr42ej32pLablPPPCW2WXY+sduWff89UPf/PebclS5YsWbJkaQ7L7N9/NPv3X839/Wezf//d7O8/mP39F7O//2T299/M/v6j9f6r9f6zid9//w+ijyJX",
}


def _ftp_icon_rgba(size):
    """내장된 초록 바탕 + 흰색 'FTP' 아이콘 (RGBA 바이트). 가장 가까운 크기를 돌려준다."""
    import zlib, base64
    key = min(_ICON_DATA, key=lambda k: abs(k - size))
    raw = zlib.decompress(base64.b64decode(_ICON_DATA[key]))
    if key != size:                                  # 필요한 크기로 최근접 보간
        out = bytearray()
        for y in range(size):
            sy = min(key - 1, y * key // size)
            for x in range(size):
                sx = min(key - 1, x * key // size)
                out += raw[(sy * key + sx) * 4:(sy * key + sx) * 4 + 4]
        raw = bytes(out)
    return raw


def _ftp_icon_png(size):
    import zlib, base64
    raw = _ftp_icon_rgba(size)
    rows = b"".join(b"\x00" + raw[y * size * 4:(y + 1) * size * 4] for y in range(size))

    def chunk(t, d):
        c = struct.pack(">I", len(d)) + t + d
        return c + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    png = (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0))
           + chunk(b"IDAT", zlib.compress(rows, 9)) + chunk(b"IEND", b""))
    return base64.b64encode(png)


def _make_icon(size=32):
    """FTP 아이콘을 HICON 으로 만들어 돌려준다."""
    raw = _ftp_icon_rgba(size)
    bmi = _BITMAPINFOHEADER(ctypes.sizeof(_BITMAPINFOHEADER), size, -size, 1, 32, 0, 0, 0, 0, 0, 0)
    bits = ctypes.c_void_p()
    color = _gdi32.CreateDIBSection(None, ctypes.byref(bmi), 0, ctypes.byref(bits), None, 0)
    buf = (ctypes.c_uint32 * (size * size)).from_address(bits.value)
    for n in range(size * size):
        r, g, b, a = raw[n * 4:n * 4 + 4]
        buf[n] = (a << 24) | (r << 16) | (g << 8) | b
    mask = _gdi32.CreateBitmap(size, size, 1, 1, None)
    info = _ICONINFO(True, 0, 0, mask, color)
    icon = _user32.CreateIconIndirect(ctypes.byref(info))
    _gdi32.DeleteObject(color)
    _gdi32.DeleteObject(mask)
    return icon


class TrayIcon:
    """시계 옆 알림 영역 아이콘. 별도 스레드에서 Win32 메시지를 처리하고,
    눌린 명령은 events 큐로 넘긴다 (tkinter 는 메인 스레드에서만 다룰 수 있으므로)."""

    MENU_TOGGLE, MENU_TOPMOST, MENU_FIX, MENU_SETTINGS, MENU_QUIT, MENU_TRANSFER = 1, 2, 4, 3, 9, 5

    def __init__(self, events, tip, states, hotkey=DEFAULT_HOTKEY, hotkey2=None):
        self.events = events                 # queue.Queue: "toggle", "topmost", "settings", "quit", ("error", msg)
        self.tip = tip
        self.states = states                 # 메뉴 체크 표시용: () -> (항상 위, 고정)
        self.hotkey = hotkey
        self.hotkey2 = hotkey2               # 파일 전송 창 단축키 (없으면 등록 안 함)
        self.hwnd = None
        self._ready = threading.Event()
        self._thread = threading.Thread(target=self._run, name="tray", daemon=True)

    def start(self):
        self._thread.start()
        self._ready.wait(3)
        return self.hwnd is not None

    def stop(self):
        if self.hwnd:
            _user32.PostMessageW(self.hwnd, _WM_CLOSE, 0, 0)
            self._thread.join(2)

    def set_hotkey(self, text):
        self.hotkey = text
        if self.hwnd:
            _user32.PostMessageW(self.hwnd, _WM_SETKEY, 0, 0)       # 등록은 창을 만든 스레드에서 해야 한다

    def _register_hotkey(self):
        for hid, text in ((_HOTKEY_ID, self.hotkey), (_HOTKEY_ID2, self.hotkey2)):
            _user32.UnregisterHotKey(self.hwnd, hid)
            if not text:
                continue
            p = parse_hotkey(text)
            if not p or not _user32.RegisterHotKey(self.hwnd, hid, p[0] | _MOD_NOREPEAT, p[1]):
                self.events.put(("error", f"{text} 단축키를 등록하지 못했습니다.\n"
                                          "다른 프로그램이 이미 쓰고 있을 수 있습니다. 설정에서 다른 키를 고르세요."))

    def set_hotkey2(self, text):
        self.hotkey2 = text
        if self.hwnd:
            _user32.PostMessageW(self.hwnd, _WM_SETKEY, 0, 0)

    def set_tip(self, text):
        self.tip = text
        if self.hwnd:
            nid = self._nid(_NIF_TIP)
            _shell32.Shell_NotifyIconW(_NIM_MODIFY, ctypes.byref(nid))

    def _nid(self, flags):
        nid = _NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(_NOTIFYICONDATAW)
        nid.hWnd = self.hwnd
        nid.uID = 1
        nid.uFlags = flags
        nid.uCallbackMessage = _WM_TRAY
        nid.hIcon = self._icon
        nid.szTip = self.tip[:127]
        return nid

    def _add_icon(self):
        nid = self._nid(_NIF_MESSAGE | _NIF_ICON | _NIF_TIP)
        return _shell32.Shell_NotifyIconW(_NIM_ADD, ctypes.byref(nid))

    def _show_menu(self):
        menu = _user32.CreatePopupMenu()
        _user32.AppendMenuW(menu, _MF_STRING, self.MENU_TOGGLE, "보이기 / 숨기기\t" + self.hotkey)
        _user32.AppendMenuW(menu, _MF_STRING, self.MENU_SETTINGS, "설정...")
        _user32.AppendMenuW(menu, _MF_SEPARATOR, 0, None)
        _user32.AppendMenuW(menu, _MF_STRING, self.MENU_QUIT, "종료")
        pt = wintypes.POINT()
        _user32.GetCursorPos(ctypes.byref(pt))
        _user32.SetForegroundWindow(self.hwnd)       # 메뉴 바깥을 누르면 닫히도록
        cmd = _user32.TrackPopupMenu(menu, _TPM_RIGHTBUTTON | _TPM_RETURNCMD | _TPM_BOTTOMALIGN,
                                     pt.x, pt.y, 0, self.hwnd, None)
        _user32.PostMessageW(self.hwnd, _WM_NULL, 0, 0)
        _user32.DestroyMenu(menu)
        name = {self.MENU_TOGGLE: "toggle", self.MENU_SETTINGS: "settings", self.MENU_QUIT: "quit"}.get(cmd)
        if name:
            self.events.put(name)

    def _wndproc(self, hwnd, msg, wparam, lparam):
        if msg == _WM_TRAY:
            if lparam == _WM_RBUTTONUP:
                self._show_menu()
            elif lparam in (_WM_LBUTTONUP, _WM_LBUTTONDBLCLK):
                self.events.put("toggle")
            return 0
        if msg == _WM_SETKEY:
            self._register_hotkey()
            return 0
        if msg == _WM_HOTKEY and wparam in (_HOTKEY_ID, _HOTKEY_ID2):
            self.events.put("toggle" if wparam == _HOTKEY_ID else "transfer")
            return 0
        if msg == _WM_SHOW:                     # 프로그램을 한 번 더 실행하면 기존 창을 보여준다
            self.events.put("show")
            return 0
        if msg == self._taskbar_created:        # 탐색기가 다시 시작되면 아이콘을 다시 올린다
            self._add_icon()
            return 0
        if msg == _WM_CLOSE:
            _shell32.Shell_NotifyIconW(_NIM_DELETE, ctypes.byref(self._nid(0)))
            _user32.UnregisterHotKey(hwnd, _HOTKEY_ID)
            _user32.UnregisterHotKey(hwnd, _HOTKEY_ID2)
            _user32.DestroyWindow(hwnd)
            return 0
        if msg == _WM_DESTROY:
            _user32.PostQuitMessage(0)
            return 0
        return _user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    def _run(self):
        try:
            hinst = _kernel32.GetModuleHandleW(None)
            self._proc = _WNDPROC(self._wndproc)          # 참조를 붙잡아 둬야 GC 되지 않는다
            wc = _WNDCLASSW()
            wc.lpfnWndProc = self._proc
            wc.hInstance = hinst
            wc.lpszClassName = _TRAY_CLASS
            _user32.RegisterClassW(ctypes.byref(wc))
            self._taskbar_created = _user32.RegisterWindowMessageW("TaskbarCreated")
            self._icon = _make_icon()
            self.hwnd = _user32.CreateWindowExW(0, wc.lpszClassName, "FtpLiteTray", 0,
                                                0, 0, 0, 0, None, None, hinst, None)
            if not self.hwnd or not self._add_icon():
                self.hwnd = None
                self.events.put(("error", "트레이 아이콘을 만들지 못했습니다."))
                return
            self._register_hotkey()
        finally:
            self._ready.set()
        msg = _MSG()
        while _user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            _user32.TranslateMessage(ctypes.byref(msg))
            _user32.DispatchMessageW(ctypes.byref(msg))
        _user32.DestroyIcon(self._icon)


# --------------------------------------------------------------------------
# 윈도우 시작 시 자동 실행 (현재 사용자 시작 프로그램 레지스트리)
# --------------------------------------------------------------------------
AUTOSTART_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
AUTOSTART_NAME = "FTP Lite"


def _autostart_command():
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    exe = Path(sys.executable)
    if exe.name.lower() == "python.exe" and (exe.parent / "pythonw.exe").exists():
        exe = exe.parent / "pythonw.exe"              # 콘솔 창이 뜨지 않게
    return f'"{exe}" "{Path(__file__).resolve()}"'


def autostart_get():
    if sys.platform != "win32":
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTOSTART_KEY) as k:
            winreg.QueryValueEx(k, AUTOSTART_NAME)
        return True
    except OSError:
        pass
    try:                                    # 옛 이름("NC Monitor")으로 등록돼 있으면 새 이름으로 옮긴다
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTOSTART_KEY, 0, winreg.KEY_READ | winreg.KEY_SET_VALUE) as k:
            winreg.QueryValueEx(k, "NC Monitor")
            winreg.SetValueEx(k, AUTOSTART_NAME, 0, winreg.REG_SZ, _autostart_command())
            winreg.DeleteValue(k, "NC Monitor")
        return True
    except OSError:
        return False


def autostart_set(on):
    if sys.platform != "win32":
        return
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTOSTART_KEY, 0, winreg.KEY_SET_VALUE) as k:
            if on:
                winreg.SetValueEx(k, AUTOSTART_NAME, 0, winreg.REG_SZ, _autostart_command())
            else:
                try:
                    winreg.DeleteValue(k, AUTOSTART_NAME)
                except FileNotFoundError:
                    pass
    except OSError:
        pass


APP_VERSION = "1.0.0"


# --------------------------------------------------------------------------
DEFAULT_RECONNECT_SEC = 10     # 연결이 끊기면 이 간격(초)마다 다시 연결
DEFAULT_KEEPALIVE_SEC = 60     # 연결된 동안 살아있는지 확인하는 간격(초)

# 상태별 (배경색, 글자색)
def cfg_seconds(cfg, key, default, low=1, high=3600):
    try:
        v = int(float(cfg.get(key, default)))
    except (TypeError, ValueError):
        return default
    return max(low, min(high, v))


# ==========================================================================
# 파일 전송 창 (M-Link 와 같은 구성)
#   위쪽 : 설비들의 FTP 데이터 서버 (가로로 나란히)
#   아래쪽: 내 PC (왼쪽 폴더 트리 + 오른쪽 폴더 목록), 전송 기록은 중간
#   제목표시줄·작업표시줄 없음, 트레이 메뉴 / 단축키(기본 Ctrl+Alt+F)로 보이기·숨기기
#   숨기면 장비와의 FTP 연결을 끊고, 다시 보이면 연결한다.
# ==========================================================================
DEFAULT_TW_HOTKEY = "Ctrl+Alt+F"

# M-Link 테마표 (라이트) - 글자 크기 9pt / 이름 10pt 굵게 / 갱신 시각 8pt, 모서리 패널 8px · 버튼·입력칸 6px
ML_LIGHT = dict(bg="#eef1f5", panel="#ffffff", panel2="#f5f7fa", panel3="#e9edf3", line="#dde3ea", line2="#c5cedb",
                text="#000000", muted="#333333", dim="#4d4d4d", blue="#2f6fe4", blue2="#1f5ccc", green="#22b35b",
                gray="#b4bdc9", yellow="#c98a00", red="#d84a4a", sel="#cce8ff", brand="#16324f", alt="#f1f2f4",
                on_accent="#ffffff", scroll_hover="#8e949e", run="#fff1b8")
ML_DARK = dict(bg="#1b1e23", panel="#262a31", panel2="#2e333b", panel3="#22262c", line="#383d46", line2="#4a505b",
               text="#e8eaed", muted="#c3c8d0", dim="#9aa1ac", blue="#5b9dff", blue2="#4a8af0", green="#3ccf7a",
               gray="#6b7380", yellow="#e0a800", red="#ff6b6b", sel="#2f4b73", brand="#9cc3ff", alt="#2b2f37",
               on_accent="#ffffff", scroll_hover="#6b7380", run="#5c4b10")
ML = dict(ML_LIGHT)
DARK_MODE = [False]


def apply_theme(dark):
    """파일 전송 창 색을 라이트/다크로 바꾼다 (창은 새로 그려야 반영됨)."""
    DARK_MODE[0] = bool(dark)
    ML.clear()
    ML.update(ML_DARK if dark else ML_LIGHT)
F_BASE, F_NAME, F_SMALL = 9, 10, 8
F_LIST = 9                      # 파일 목록 글자 (조금 크게)


def work_area(win):
    """작업표시줄을 뺀 화면 영역 [x, y, w, h]."""
    if IS_WINDOWS:
        try:
            r = wintypes.RECT()
            if ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(r), 0):
                return [r.left, r.top, r.right - r.left, r.bottom - r.top]
        except Exception:
            pass
    return [0, 0, win.winfo_screenwidth(), win.winfo_screenheight()]


def tw_hotkey_cfg(cfg):
    h = str(cfg.get("tw_hotkey", DEFAULT_TW_HOTKEY))
    return h if parse_hotkey(h) else DEFAULT_TW_HOTKEY


def pc_default_dir(cfg):
    """내 PC 쪽 시작 폴더 (설정 값 → D:\\ → C:\\)."""
    d = str(cfg.get("pc_dir", "") or "").strip()
    if d and Path(d).is_dir():
        return Path(d)
    if IS_WINDOWS:
        return Path("D:\\") if Path("D:\\").exists() else Path("C:\\")
    return Path.home()


def ncdata_dir(cfg):
    """내 PC 쪽 NCDATA 폴더 (설정 값, 없으면 D:\\NCDATA > C:\\NCDATA)."""
    d = str(cfg.get("ncdata_dir", "") or "").strip()
    if d:
        return Path(d)
    if IS_WINDOWS:
        for drv in ("D", "C"):
            if Path(f"{drv}:\\NCDATA").is_dir():
                return Path(f"{drv}:\\NCDATA")
        return Path("D:\\NCDATA") if Path("D:\\").exists() else Path("C:\\NCDATA")
    return Path.home() / "NCDATA"


# ---------------------------------------------------------------- FTP 목록 해석 / 오류 문구
@dataclass
class ListEntry:
    name: str
    is_dir: bool
    size: Optional[int]
    mtime: Optional[object]


_MONTHS = {m: i + 1 for i, m in enumerate(("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct",
                                            "Nov", "Dec"))}
_UNIX_RE = re.compile(r"^(?P<type>[-dl])[-rwxstST]{9}\+?\s+\d+\s+\S+\s+\S+\s+(?P<size>\d+)\s+(?P<mon>[A-Za-z]{3})"
                      r"\s+(?P<day>\d{1,2})\s+(?P<yt>\d{1,2}:\d{2}|\d{4})\s+(?P<name>.+)$")
_DOS_RE = re.compile(r"^(?P<mon>\d{2})-(?P<day>\d{2})-(?P<year>\d{2,4})\s+(?P<hour>\d{1,2}):(?P<minute>\d{2})\s*"
                     r"(?P<ampm>AM|PM)?\s+(?P<sz><DIR>|\d+)\s+(?P<name>.+)$", re.I)


def parse_list_line(line, now=None):
    """FTP LIST 한 줄 -> ListEntry (유닉스식·DOS식 둘 다). 해석할 수 없으면 None."""
    from datetime import datetime
    line = line.rstrip("\r\n")
    now = now or datetime.now()
    m = _UNIX_RE.match(line)
    if m:
        mon = _MONTHS.get(m["mon"].title())
        mt = None
        try:
            if ":" in m["yt"]:
                hh, mm = m["yt"].split(":")
                mt = datetime(now.year, mon, int(m["day"]), int(hh), int(mm))
                if mt > now:
                    mt = mt.replace(year=now.year - 1)
            else:
                mt = datetime(int(m["yt"]), mon, int(m["day"]))
        except (ValueError, TypeError):
            mt = None
        return ListEntry(m["name"], m["type"] == "d", int(m["size"]), mt)
    m = _DOS_RE.match(line.strip())
    if m:
        try:
            year = int(m["year"])
            year += 2000 if year < 70 else (1900 if year < 100 else 0)
            hour = int(m["hour"]) % 12 + (12 if (m["ampm"] or "").upper() == "PM" else 0) if m["ampm"] \
                else int(m["hour"])
            mt = datetime(year, int(m["mon"]), int(m["day"]), hour, int(m["minute"]))
        except ValueError:
            mt = None
        isdir = m["sz"].upper() == "<DIR>"
        return ListEntry(m["name"], isdir, None if isdir else int(m["sz"]), mt)
    return None


_FTP_CODES = {
    "421": "장비가 지금 연결을 받지 않습니다 (접속 수 초과일 수 있음)", "425": "데이터 연결을 열지 못했습니다",
    "426": "전송이 중간에 끊겼습니다", "450": "파일을 지금 쓸 수 없습니다 (사용 중일 수 있음)",
    "530": "로그인 실패 · 사용자/비밀번호를 확인하세요", "550": "파일을 찾을 수 없거나 권한이 없습니다",
    "552": "장비 저장 공간이 부족합니다", "553": "파일 이름을 쓸 수 없습니다",
}


def ftp_error_text(e):
    if isinstance(e, (socket.timeout, TimeoutError)):
        return "응답 없음 — 장비가 꺼졌거나 네트워크를 확인하세요"
    if isinstance(e, ConnectionRefusedError):
        return "연결이 거부됐습니다 — 장비 FTP 서비스를 확인하세요"
    if isinstance(e, (ConnectionResetError, BrokenPipeError, EOFError)):
        return "연결이 끊겼습니다 — 장비가 재부팅됐거나 케이블을 확인하세요"
    if isinstance(e, UnicodeError):
        return "파일 이름을 읽지 못했습니다"
    s = str(e)
    if s[:3] in _FTP_CODES:
        return _FTP_CODES[s[:3]]
    code = getattr(e, "winerror", None) or getattr(e, "errno", None)
    if code in (10060, 110):
        return "응답 없음 — 장비가 꺼졌거나 네트워크를 확인하세요"
    if code in (10061, 111):
        return "연결이 거부됐습니다 — 장비 FTP 서비스를 확인하세요"
    if code in (10065, 10051, 113, 101):
        return "장비에 닿을 수 없습니다 — 네트워크 경로를 확인하세요"
    if code in (10054, 10053, 104, 32):
        return "연결이 끊겼습니다 — 장비가 재부팅됐거나 케이블을 확인하세요"
    return f"통신 오류 — 장비·네트워크를 확인하세요 ({s[:60]})"


class LocalError(Exception):
    """내 PC 쪽 파일 오류 (장비 연결은 그대로 유지)."""


class FtpSession(threading.Thread):
    """장비 하나와의 FTP 연결을 맡는 스레드. 명령은 submit() 으로 넣고, 결과는 events 큐로 나온다.
    events 항목: (key, 종류, 내용)  종류 = state / list / progress / done / error / log"""

    def __init__(self, key, host, user, password, events, retry_s=10, keepalive_s=20, start_cwd="/"):
        super().__init__(name=f"ftp-{key}", daemon=True)
        self.start_cwd = start_cwd or "/"
        self.key, self.host, self.user, self.password = key, host, user, password
        self.events, self.retry_s, self.keepalive_s = events, retry_s, keepalive_s
        self.cmds = queue.Queue()
        self.halt = threading.Event()
        self.pasv = True                  # 목록이 안 열리면 능동 모드로 바꿔 다시 시도
        self._pasv_ok = False

    def post(self, kind, payload=None):
        self.events.put((self.key, kind, payload))

    def submit(self, *cmd):
        self.cmds.put(cmd)

    def stop(self):
        self.halt.set()
        self.cmds.put(None)

    # ---- 스레드 본체
    def run(self):
        ftp = None
        while not self.halt.is_set():
            if ftp is None:
                self.post("state", ("connecting", "", 0))
                try:
                    t0 = time.perf_counter()
                    ftp = self._connect()
                    self.post("state", ("on", "", round((time.perf_counter() - t0) * 1000)))
                    if self.start_cwd not in ("/", ""):
                        try:
                            ftp.cwd(self.start_cwd)              # 숨겼다 다시 열어도 보던 폴더로
                        except all_errors:
                            pass
                    self._list(ftp)
                    self._pasv_ok = True
                except (*all_errors, UnicodeError) as e:
                    if ftp is not None and not self._pasv_ok:
                        self.pasv = not self.pasv
                    ftp = self._drop(ftp)
                    if self.halt.is_set():
                        break
                    self.post("state", ("off", ftp_error_text(e), 0))
                    self._drain()
                    self.halt.wait(self.retry_s)
                    continue
            try:
                cmd = self.cmds.get(timeout=self.keepalive_s)
            except queue.Empty:
                cmd = ("noop",)
            if cmd is None or self.halt.is_set():
                break
            try:
                self._exec(ftp, cmd)
            except LocalError as e:
                self.post("error", (cmd[0], str(e)))
            except error_perm as e:
                self.post("error", (cmd[0], ftp_error_text(e)))
            except (*all_errors, UnicodeError) as e:
                self.post("error", (cmd[0], ftp_error_text(e)))
                ftp = self._drop(ftp)
            except Exception as e:                       # 예상 못 한 오류로 스레드가 조용히 죽지 않게
                self.post("error", (cmd[0], f"내부 오류: {e!r}"))
        self._drop(ftp)

    def _drain(self):
        try:
            while True:
                self.cmds.get_nowait()
        except queue.Empty:
            pass

    @staticmethod
    def _drop(ftp):
        if ftp is not None:
            try:
                ftp.close()
            except Exception:
                pass
        return None

    def _connect(self):
        ftp = FTP()
        ftp.connect(self.host, FTP_PORT, timeout=8)
        ftp.login(self.user or "anonymous", self.password)
        ftp.set_pasv(self.pasv)
        try:
            ftp.voidcmd("TYPE I")
        except all_errors:
            pass                                     # 모르는 장비는 그냥 진행
        return ftp

    def _list(self, ftp):
        lines = []
        ftp.retrlines("LIST", lines.append)
        entries = [e for e in (parse_list_line(ln) for ln in lines) if e and e.name not in (".", "..")]
        if lines and not entries and not all(ln.lower().startswith("total") for ln in lines):
            names = []
            ftp.retrlines("NLST", names.append)                 # LIST 모양을 못 읽으면 이름만이라도
            entries = [ListEntry(n.strip(), False, None, None) for n in names if n.strip()]
        try:
            cwd = ftp.pwd()
        except all_errors:
            cwd = "/"
        self.start_cwd = cwd
        self.post("list", (cwd, entries))

    # ---- 명령 처리
    def _exec(self, ftp, cmd):
        op = cmd[0]
        if op == "noop":
            ftp.voidcmd("NOOP")
        elif op == "list":
            self._list(ftp)
        elif op == "cd":
            ftp.cwd(cmd[1])
            self._list(ftp)
        elif op == "upload":
            self._upload(ftp, cmd[1])
        elif op == "download":
            self._download(ftp, cmd[1], cmd[2])
        elif op == "delete":
            self._remove(ftp, cmd[1])
        elif op == "rename":
            ftp.rename(cmd[1], cmd[2])
            self.post("log", f"{cmd[1]} → {cmd[2]} 이름 바꿈")
            self._list(ftp)

    def _progress(self, label, done, total, idx, count, last=[0.0]):
        now = time.monotonic()
        if done >= total or now - last[0] > 0.1:
            last[0] = now
            self.post("progress", (label, done, total, idx, count))

    def _upload(self, ftp, paths):
        ok = bad = 0
        for i, path in enumerate(paths, 1):
            path = Path(path)
            try:
                total = path.stat().st_size
                sent = [0]

                def cb(block, _name=path.name, _i=i):
                    sent[0] += len(block)
                    self._progress(_name, sent[0], total, _i, len(paths))
                with open(path, "rb") as f:
                    ftp.storbinary("STOR " + norm_prog_name(path.name), f, 8192, cb)
                ok += 1
                self.post("uploaded", str(path))
                self.post("log", f"{path.name} 올리기 완료")
            except OSError as e:
                if isinstance(e, (ConnectionError, socket.timeout)) or getattr(e, "errno", None) in (10054, 10053, 104, 32):
                    raise
                bad += 1
                self.post("error", ("upload", f"{path.name}: 내 PC 파일을 읽지 못했습니다 ({e.strerror or e})"))
            except (error_perm, error_temp) as e:
                bad += 1
                try:
                    ftp.delete(norm_prog_name(path.name))                # 끊긴 올리기의 반쪽 파일은 장비에서 지운다
                except all_errors:
                    pass
                self.post("error", ("upload", f"{path.name}: {ftp_error_text(e)}"))
        self.post("done", ("upload", ok, bad))
        self._list(ftp)

    def _download(self, ftp, names, local_dir):
        ok = bad = 0
        local_dir = Path(local_dir)
        for i, name in enumerate(names, 1):
            dest_name = norm_prog_name(name)
            part = local_dir / (dest_name + ".part")
            try:
                try:
                    total = ftp.size(name) or 0
                except all_errors:
                    total = 0
                got = [0]
                try:
                    f = open(part, "wb")
                except OSError as e:
                    raise LocalError(f"{name}: 내 PC 폴더에 쓰지 못했습니다 ({e.strerror or e})")
                try:
                    with f:
                        def cb(block, _name=name, _i=i):
                            f.write(block)
                            got[0] += len(block)
                            self._progress(_name, got[0], max(total, got[0]), _i, len(names))
                        ftp.retrbinary("RETR " + name, cb)
                    os.replace(part, local_dir / dest_name)
                except BaseException:
                    try:
                        part.unlink()
                    except OSError:
                        pass
                    raise
                ok += 1
                self.post("log", f"{name} 내려받기 완료")
            except LocalError as e:
                bad += 1
                self.post("error", ("download", str(e)))
            except (error_perm, error_temp) as e:
                bad += 1
                self.post("error", ("download", f"{name}: {ftp_error_text(e)}"))
            except OSError as e:
                if isinstance(e, (ConnectionError, socket.timeout)) or getattr(e, "errno", None) in (10054, 10053, 104, 32):
                    raise
                bad += 1
                self.post("error", ("download", f"{name}: 내 PC 에 쓰지 못했습니다 ({e.strerror or e})"))
        self.post("done", ("download", ok, bad))

    def _rm_dir(self, ftp, name, depth=0):
        """폴더와 그 안의 모든 파일·폴더를 지운다 (FTP 는 비어 있지 않은 폴더를 한 번에 못 지움)."""
        try:
            ftp.rmd(name)
            return
        except error_perm:
            if depth > 8:
                raise
        base = ftp.pwd()
        try:
            ftp.cwd(name)
            lines = []
            ftp.retrlines("LIST", lines.append)
            for e in (x for x in (parse_list_line(ln) for ln in lines) if x and x.name not in (".", "..")):
                if e.is_dir:
                    self._rm_dir(ftp, e.name, depth + 1)
                else:
                    ftp.delete(e.name)
        finally:
            ftp.cwd(base)
        ftp.rmd(name)

    def _remove(self, ftp, names):
        ok = bad = 0
        for name in names:
            try:
                try:
                    ftp.delete(name)
                except error_perm:
                    self._rm_dir(ftp, name)              # 폴더면 안의 파일까지 지우고 폴더 삭제
                ok += 1
                self.post("log", f"{name} 삭제")
            except (error_perm, error_temp) as e:
                bad += 1
                self.post("error", ("delete", f"{name}: {ftp_error_text(e)}"))
        self.post("done", ("delete", ok, bad))
        self._list(ftp)


# ---------------------------------------------------------------- M-Link 모양의 위젯
def _ml_label(parent, text, size=F_BASE, bold=False, fg=None, bg=None, **kw):
    return tk.Label(parent, text=text, font=ui_font(size, bold), fg=fg or ML["text"], bg=bg or ML["panel"], **kw)


def _ml_line(parent, color=None):
    return tk.Frame(parent, height=1, bg=color or ML["line"])


def _rgb(c):
    return int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)


def _mix(c1, c2, t):
    a, b = _rgb(c1), _rgb(c2)
    return "#%02x%02x%02x" % tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


def _in_poly(x, y, pts):
    inside = False
    j = len(pts) - 1
    for i in range(len(pts)):
        xi, yi = pts[i]
        xj, yj = pts[j]
        if (yi > y) != (yj > y) and x < (xj - xi) * (y - yi) / (yj - yi) + xi:
            inside = not inside
        j = i
    return inside


def _shape_test(sh):
    """도형 -> (x0, y0, x1, y1, 안에 있나(x, y), 색)"""
    kind = sh[0]
    if kind == "circle":
        _, cx, cy, r, col = sh
        return cx - r, cy - r, cx + r, cy + r, (lambda x, y: (x - cx) ** 2 + (y - cy) ** 2 <= r * r), col
    if kind == "ring":
        _, cx, cy, r, w, col = sh
        lo, hi = (r - w / 2) ** 2, (r + w / 2) ** 2
        return cx - r - w, cy - r - w, cx + r + w, cy + r + w, \
            (lambda x, y: lo <= (x - cx) ** 2 + (y - cy) ** 2 <= hi), col
    if kind == "line":
        _, x1, y1, x2, y2, w, col = sh
        dx, dy = x2 - x1, y2 - y1
        L2 = dx * dx + dy * dy or 1e-9
        hw2 = (w / 2) ** 2

        def hit(x, y):
            t = max(0.0, min(1.0, ((x - x1) * dx + (y - y1) * dy) / L2))
            return (x - x1 - t * dx) ** 2 + (y - y1 - t * dy) ** 2 <= hw2
        return min(x1, x2) - w, min(y1, y2) - w, max(x1, x2) + w, max(y1, y2) + w, hit, col
    _, pts, col = sh                                   # poly
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    return min(xs), min(ys), max(xs), max(ys), (lambda x, y: _in_poly(x, y, pts)), col


def render_icon(master, shapes, W, H, bg, S=4):
    """도형들을 S x S 겹샘플링으로 그려 가장자리가 부드러운 PhotoImage 로 만든다 (캔버스 선은 계단지기 때문)."""
    prep = [_shape_test(sh) for sh in shapes]
    bgc = _rgb(bg)
    offs = [(i + 0.5) / S for i in range(S)]
    rows = []
    for py in range(H):
        row = []
        for px in range(W):
            cand = [(t, _rgb(c)) for (x0, y0, x1, y1, t, c) in prep if x0 - 1 <= px <= x1 + 1 and y0 - 1 <= py <= y1 + 1]
            if not cand:
                row.append(bg)
                continue
            r = g = b = 0
            for oy in offs:
                for ox in offs:
                    col = bgc
                    for test, c in cand:
                        if test(px + ox, py + oy):
                            col = c
                    r += col[0]
                    g += col[1]
                    b += col[2]
            n = S * S
            row.append("#%02x%02x%02x" % (r // n, g // n, b // n))
        rows.append("{" + " ".join(row) + "}")
    img = tk.PhotoImage(master=master, width=W, height=H)
    img.put(" ".join(rows))
    return img


class PillToggle(tk.Canvas):
    """알약 모양 켜고 끄는 버튼: 휴지통 아이콘 + 글자. 켜면 빨간 테두리·연한 빨강 배경·굵은 글씨."""
    RED = "#ff3b30"

    def __init__(self, master, text, variable, command=None):
        self.var, self._cmd, self._text = variable, command, text
        self._f_off, self._f_on = ui_font(F_BASE, False), ui_font(F_BASE, True)
        wid = max(tkfont.Font(font=self._f_off).measure(text), tkfont.Font(font=self._f_on).measure(text))
        self.W, self.H = int(wid) + 46, 28
        super().__init__(master, width=self.W, height=self.H, bg=ML["panel"], highlightthickness=0, bd=0, cursor="hand2")
        self.bind("<Button-1>", self._click)
        self.var.trace_add("write", lambda *_: self._draw())
        self._draw()

    def _click(self, _e=None):
        self.var.set(not self.var.get())
        if self._cmd:
            self._cmd()

    def _draw(self):
        c = self
        try:
            c.delete("all")
        except tk.TclError:
            return
        on = bool(self.var.get())
        edge = self.RED if on else ML["dim"]
        ink = self.RED if on else ML["muted"]
        fill = _mix(ML["panel"], self.RED, 0.14 if not DARK_MODE[0] else 0.22) if on else ML["panel"]
        c.create_polygon(rr_points(1, 1, self.W - 1, self.H - 1, 13), smooth=True, fill=fill, outline=edge, width=1.4)
        x, y = 17, self.H / 2                                    # 휴지통 아이콘
        c.create_line(x - 5, y - 4, x + 5, y - 4, fill=ink, width=1.6, capstyle="round")
        c.create_line(x - 1.5, y - 6, x + 1.5, y - 6, fill=ink, width=1.6, capstyle="round")
        c.create_polygon(x - 4, y - 2, x + 4, y - 2, x + 3.2, y + 6, x - 3.2, y + 6, fill="", outline=ink, width=1.4)
        c.create_text(30, y, text=self._text, anchor="w", font=self._f_on if on else self._f_off, fill=ink)


class IconBtn(tk.Canvas):
    """맥OS 느낌 그림 버튼 (부드러운 가장자리로 직접 그림).
    close(빨강) / full·restore(노랑) / moon·sun(회색, 다크모드) / gear(회색 원 테두리 + 톱니)."""

    CACHE = {}
    COLORS = {            # 채움, 테두리, 안쪽 그림 색
        "close": ("#ff5f57", "#e0443e", "#7d0a0a"),
        "full": ("#febc2e", "#dea123", "#7a4f00"),
        "restore": ("#febc2e", "#dea123", "#7a4f00"),
        "moon": ("#a9afb7", "#8a9098", "#3d4249"),
        "sun": ("#a9afb7", "#8a9098", "#3d4249"),
    }

    def __init__(self, master, icon, command, bg=None, size=36):
        self._bg = bg or ML["panel"]
        self.W = self.H = size
        super().__init__(master, width=self.W, height=self.H, bg=self._bg, highlightthickness=0, bd=0, cursor="hand2")
        self.icon, self._command = icon, command
        self._hover = self._down = False
        self._img = None
        self.bind("<Enter>", lambda _e: self._set(hover=True))
        self.bind("<Leave>", lambda _e: self._set(hover=False, down=False))
        self.bind("<ButtonPress-1>", lambda _e: self._set(down=True))
        self.bind("<ButtonRelease-1>", self._release)
        self._draw()

    def set_icon(self, icon):
        self.icon = icon
        self._draw()

    def _set(self, hover=None, down=None):
        if hover is not None:
            self._hover = hover
        if down is not None:
            self._down = down
        self._draw()

    def _release(self, e):
        fire = self._down and 0 <= e.x <= self.W and 0 <= e.y <= self.H
        self._down = False
        self._draw()
        if fire:
            self._command()

    def _shapes(self):
        import math
        cx = cy = self.W / 2
        st = 2 if self._down else (1 if self._hover else 0)
        if self.icon == "broom":                      # 빗자루: 자루 + 부채꼴 솔
            col = ("#8b9199", "#c3c8cf" if DARK_MODE[0] else "#4e545c", "#6f757d")[st]
            k = self.W / 28.0
            P = lambda x, y: (x * k, y * k)
            return [("line", *P(22.5, 4.5), *P(14.5, 16.5), 2.4 * k, col),
                    ("poly", [P(11.5, 15.3), P(17, 18.8), P(15.2, 27), P(4, 20)], col),
                    ("line", *P(13.5, 18.2), *P(9.2, 24.2), 0.9 * k, self._bg),
                    ("line", *P(15.2, 19.8), *P(12, 25.2), 0.9 * k, self._bg)]
        if self.icon == "gear":
            col = ("#9aa0a8", "#c3c8cf" if DARK_MODE[0] else "#6b7078", "#8b9199")[st]
            sh = [("ring", cx, cy, 12.4, 1.7, col)]
            pts = []
            for k in range(8):
                a = k * math.pi / 4
                for ro, da in ((6.0, -0.36), (8.4, -0.21), (8.4, 0.21), (6.0, 0.36)):
                    pts.append((cx + ro * math.cos(a + da), cy + ro * math.sin(a + da)))
            sh += [("poly", pts, col), ("circle", cx, cy, 3.0, self._bg)]
            return sh
        fill, edge, ink = self.COLORS[self.icon]
        shade = (0, 0.10, 0.22)[st]
        fill = _mix(fill, ink, shade)
        sh = [("circle", cx, cy, 12.4, edge), ("circle", cx, cy, 11.4, fill)]
        k = dict(w=2.4, col=ink)
        if self.icon == "close":
            sh += [("line", cx - 4.2, cy - 4.2, cx + 4.2, cy + 4.2, 2.4, ink),
                   ("line", cx - 4.2, cy + 4.2, cx + 4.2, cy - 4.2, 2.4, ink)]
        elif self.icon == "full":                     # 바깥쪽을 향한 삼각형 한 쌍
            sh += [("poly", [(cx + 5, cy - 5), (cx + 5, cy + 1.2), (cx - 1.2, cy - 5)], ink),
                   ("poly", [(cx - 5, cy + 5), (cx - 5, cy - 1.2), (cx + 1.2, cy + 5)], ink)]
        elif self.icon == "restore":                  # 안쪽을 향한 삼각형 한 쌍
            sh += [("poly", [(cx + 5, cy + 0.6), (cx - 0.6, cy + 0.6), (cx - 0.6, cy - 5)], ink),
                   ("poly", [(cx - 5, cy - 0.6), (cx + 0.6, cy - 0.6), (cx + 0.6, cy + 5)], ink)]
        elif self.icon == "moon":
            sh += [("circle", cx - 0.3, cy + 0.3, 5.2, ink), ("circle", cx + 2.5, cy - 2.2, 4.4, fill)]
        else:                                         # sun
            for n in range(8):
                a = n * math.pi / 4
                sh.append(("line", cx + 5.0 * math.cos(a), cy + 5.0 * math.sin(a),
                           cx + 7.0 * math.cos(a), cy + 7.0 * math.sin(a), 1.5, ink))
            sh.append(("circle", cx, cy, 3.4, ink))
        return sh

    def _draw(self):
        key = (self.icon, self._hover, self._down, self._bg, DARK_MODE[0], self.W)
        img = IconBtn.CACHE.get(key)
        if img is None:
            img = IconBtn.CACHE[key] = render_icon(self, self._shapes(), self.W, self.H, self._bg)
        try:
            self.delete("all")
            self._img = img
            self.create_image(0, 0, image=img, anchor="nw")
        except tk.TclError:
            pass


class MButton(tk.Canvas):
    """QPushButton: 흰 바탕 + 1px 테두리, 모서리 6px. 마우스를 올리면 글자·테두리 파랑, 누르면 panel3."""

    def __init__(self, master, text, command=None, kind="default", width=None, bg=None):
        self._font = tkfont.Font(family=UI_FONT, size=F_BASE)
        self._text, self._command, self._kind = text, command, kind
        self._bw = width or self._font.measure(text) + 24
        self._bh = self._font.metrics("linespace") + 10
        super().__init__(master, width=self._bw, height=self._bh, bg=bg or ML["panel"], highlightthickness=0, bd=0,
                         cursor="hand2")
        self._hover = self._down = False
        self._enabled = True
        self.bind("<Enter>", lambda _e: self._set(hover=True))
        self.bind("<Leave>", lambda _e: self._set(hover=False, down=False))
        self.bind("<ButtonPress-1>", lambda _e: self._set(down=True))
        self.bind("<ButtonRelease-1>", self._release)
        self._draw()

    def _set(self, hover=None, down=None):
        if hover is not None:
            self._hover = hover
        if down is not None:
            self._down = down
        self._draw()

    def _release(self, e):
        fire = self._down and self._enabled and 0 <= e.x <= self._bw and 0 <= e.y <= self._bh
        self._down = False
        self._draw()
        if fire and self._command:
            self._command()

    def _draw(self):
        try:
            self.delete("all")
            if self._kind == "primary":
                fill = edge = ML["blue2"] if self._hover else ML["blue"]
                fg = ML["on_accent"]
            else:
                fill = ML["panel3"] if self._down else (ML["panel2"] if self._hover else ML["panel"])
                edge = ML["blue2"] if self._down else (ML["blue"] if self._hover else ML["line2"])
                fg = ML["red"] if (self._kind == "danger" and self._hover) else (ML["blue"] if self._hover else ML["text"])
            if not self._enabled:
                fill, edge, fg = ML["panel"], ML["line"], ML["gray"]
            self.create_polygon(rr_points(0, 0, self._bw - 1, self._bh - 1, 6), smooth=True, fill=fill, outline=edge)
            self.create_text(self._bw / 2, self._bh / 2 + (1 if self._down else 0), text=self._text,
                             font=self._font, fill=fg)
        except tk.TclError:
            pass

    def set_text(self, text):
        self._text = text
        self._draw()

    def set_enabled(self, on):
        self._enabled = bool(on)
        self.configure(cursor="hand2" if on else "arrow")
        self._draw()


class MEntry(tk.Canvas):
    """QLineEdit: panel2 바탕, 1px 테두리, 모서리 6px, 입력 중엔 테두리 파랑. Enter 를 누르면 on_enter."""

    def __init__(self, master, var, on_enter=None, height=28):
        self._bh = height
        super().__init__(master, height=height, bg=ML["panel"], highlightthickness=0, bd=0)
        self._focus = False
        self.entry = tk.Entry(self, textvariable=var, font=ui_font(F_BASE), bd=0, relief="flat", bg=ML["panel2"],
                              fg=ML["text"], insertbackground=ML["text"], highlightthickness=0)
        self._win = self.create_window(8, height // 2, window=self.entry, anchor="w")
        self.entry.bind("<FocusIn>", lambda _e: self._focus_set(True))
        self.entry.bind("<FocusOut>", lambda _e: self._focus_set(False))
        self.entry.bind("<Button-1>", lambda _e: self.entry.focus_force())
        if on_enter:
            self.entry.bind("<Return>", lambda _e: on_enter())
        self.bind("<Configure>", self._on_size)

    def _focus_set(self, on):
        self._focus = on
        self._on_size()

    def _on_size(self, _e=None):
        w = self.winfo_width()
        self.delete("edge")
        self.create_polygon(rr_points(0, 0, w - 1, self._bh - 1, 6), smooth=True, fill=ML["panel2"],
                            outline=ML["blue"] if self._focus else ML["line2"], tags="edge")
        self.tag_lower("edge")
        self.itemconfigure(self._win, width=max(1, w - 16))


class MPanel(RoundedCard):
    """QFrame#Panel: 흰 바탕 + 1px 테두리(line2) + 모서리 8px."""

    def __init__(self, master):
        super().__init__(master, outer=ML["bg"], fill=ML["panel"], radius=8, pad=1, expand=True)

    def _on_canvas(self, e):
        self.delete("bg")
        self.create_polygon(rr_points(0, 0, e.width - 1, e.height - 1, 8), smooth=True, fill=ML["panel"],
                            outline=ML["line2"], tags="bg")
        self.tag_lower("bg")
        self.itemconfigure(self._win, width=max(1, e.width - 2), height=max(1, e.height - 2))


def ml_styles(root):
    st = ttk.Style(root)
    st.configure("MLT.Treeview", rowheight=32, font=ui_font(F_LIST), background=ML["panel"],
                 fieldbackground=ML["panel"], foreground=ML["text"], borderwidth=0, relief="flat", indent=16)
    st.map("MLT.Treeview", background=[("selected", ML["sel"])], foreground=[("selected", ML["text"])])
    st.layout("MLT.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
    st.configure("ML.Treeview", rowheight=28, font=ui_font(F_LIST), background=ML["panel"],
                 fieldbackground=ML["panel"], foreground=ML["text"], borderwidth=0, relief="flat")
    st.map("ML.Treeview", background=[("selected", ML["sel"])], foreground=[("selected", ML["text"])])
    st.layout("ML.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
    st.configure("ML.Treeview.Heading", background=ML["panel2"], foreground=ML["muted"], font=ui_font(F_BASE, True),
                 relief="flat", borderwidth=0, padding=(8, 4))
    st.map("ML.Treeview.Heading", background=[("active", ML["panel3"])])
    st.configure("ML.TCombobox", fieldbackground=ML["panel"], background=ML["panel"], foreground=ML["text"],
                 bordercolor=ML["line2"], lightcolor=ML["line2"], darkcolor=ML["line2"], arrowcolor=ML["dim"],
                 padding=(6, 2), relief="flat")
    st.map("ML.TCombobox", fieldbackground=[("readonly", ML["panel"])], bordercolor=[("focus", ML["blue"])])
    root.option_add("*TCombobox*Listbox.font", ui_font(F_LIST))
    root.option_add("*TCombobox*Listbox.selectBackground", ML["sel"])
    root.option_add("*TCombobox*Listbox.selectForeground", ML["text"])
    st.configure("ML.Horizontal.TProgressbar", troughcolor=ML["panel"], background=ML["blue"],
                 bordercolor=ML["line"], lightcolor=ML["blue"], darkcolor=ML["blue"], thickness=10)
    st.configure("ML.Vertical.TScrollbar", background=ML["line2"], troughcolor=ML["panel"],
                 bordercolor=ML["panel"], lightcolor=ML["line2"], darkcolor=ML["line2"], arrowcolor=ML["panel"],
                 relief="flat", width=10)
    st.map("ML.Vertical.TScrollbar", background=[("active", ML["scroll_hover"]), ("pressed", ML["blue"])])
    st.layout("ML.Vertical.TScrollbar", [("Vertical.Scrollbar.trough", {"sticky": "ns", "children": [
        ("Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])


def norm_prog_name(name):
    """o12 -> O0012, o2.tap -> O0002.tap. 숫자가 5자리 이상(O00000~)이면 그대로 둔다."""
    m = re.fullmatch(r"[oO](\d{1,4})(\..*)?", name)
    return f"O{m.group(1).zfill(4)}{m.group(2) or ''}" if m else name


def _fmt_size(n):
    if n is None:
        return ""
    return f"{(n + 1023) // 1024:,}KB" if n else "0KB"


def _fmt_time(dt):
    """26.10.1 10:14"""
    return f"{dt.year % 100}.{dt.month}.{dt.day} {dt:%H:%M}" if dt else ""


PC_COLS = (("name", "이름", 90, "w", True), ("size", "크기", 66, "e", False),
           ("mtime", "최종 수정", 112, "center", False))


def _file_list(parent, on_open, cols=None):
    """이름 / 수정 시각 / 용량 / 종류 목록 (흰 배경 한 가지, 스크롤바 없음 - 마우스 휠로 넘김).
    (바깥 프레임, Treeview, '빈 폴더' 라벨) 을 돌려준다."""
    wrap = tk.Frame(parent, bg=ML["panel"])
    cols = cols or (("name", "이름", 230, "w", True), ("mtime", "최종 수정", 140, "center", False),
                    ("size", "용량", 90, "e", False), ("kind", "종류", 80, "w", False))
    tree = ttk.Treeview(wrap, columns=[c[0] for c in cols], show="headings", selectmode="extended",
                        style="ML.Treeview")
    for key, text, width, anchor, stretch in cols:
        tree.heading(key, text=text, anchor="center" if key in ("size", "mtime") else anchor)
        tree.column(key, width=width, anchor=anchor, stretch=stretch)
    tree.tag_configure("alt", background=ML["alt"])           # 줄마다 투명(흰색) - 회색 번갈아
    tree.tag_configure("dir", font=ui_font(F_LIST, True))
    tree.tag_configure("run", background=ML["run"])      # 장비가 지금 돌리는 프로그램
    tree.pack(fill="both", expand=True)
    empty = tk.Label(wrap, text="빈 폴더", font=ui_font(F_BASE), bg=ML["panel"], fg=ML["dim"])
    tree.bind("<Double-1>", lambda e: on_open(tree.identify_row(e.y)))
    tree.bind("<Return>", lambda _e: on_open(tree.focus()))
    tree.bind("<Control-a>", lambda _e: (tree.selection_set(tree.get_children()), "break")[1])
    return wrap, tree, empty


def _ask_overwrite(win, title, names):
    shown = ", ".join(names[:5]) + (" …" if len(names) > 5 else "")
    return messagebox.askyesno(title, f"이미 있는 파일 {len(names)}개가 있습니다.\n{shown}\n\n"
                                     "덮어쓸까요? (아니오 = 그 파일은 건너뜀)", parent=win)


def _prog_match(name, prog):
    """파일 이름이 프로그램 번호와 같은가 (O0123 / O123 / 0123.NC / %O123 ...)."""
    m = re.match(r"^%?[Oo]?0*(\d+)(\..*)?$", name)
    return bool(m) and int(m.group(1)) == prog


def open_in_notepad(path):
    """파일 열기: 설정에서 정한 프로그램이 있으면 그것으로, 없거나 실행에 실패하면 메모장으로."""
    ed = EDITOR_PATH[0]
    if ed and os.path.isfile(ed):
        try:
            subprocess.Popen([ed, str(path)])
            return True
        except OSError:
            pass
    try:
        if IS_WINDOWS:
            subprocess.Popen(["notepad.exe", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
        return True
    except OSError:
        return False


# ---------------------------------------------------------------- 설비 패널 (위쪽)
class MachinePanel:
    title = ""

    def __init__(self, parent, tw, key, num, name, site, retry_s):
        self.tw, self.key, self.num, self.name, self.site = tw, key, num, name, site
        self.title = name
        self.retry_s = retry_s
        self._last_click = 0.0
        self.entries = []
        self.rows = {}
        self.cwd = tw.cwd_mem.get(name, "/")           # 숨겼다 다시 열어도 보던 폴더를 기억
        self.state = "off"
        self.session = None
        self.busy = False
        self.card = MPanel(parent)
        inner = self.card.inner

        # 머리 칸: 장비명
        self.hd = hd = tk.Frame(inner, bg=ML["panel"])
        hd.pack(fill="x")
        row1 = tk.Frame(hd, bg=ML["panel"])
        row1.pack(fill="x", padx=8, pady=(6, 6))
        self.name_lbl = _ml_label(row1, name, F_NAME, True, ML["text"])
        self.name_lbl.pack(side="left", padx=(2, 8))
        _ml_line(inner, ML["line2"]).pack(fill="x")

        # 경로줄
        pr = tk.Frame(inner, bg=ML["panel"])
        pr.pack(fill="x", padx=8, pady=5)
        self.up_btn = MButton(pr, "↑", self.up, width=30)
        self.up_btn.pack(side="left", padx=(0, 6))
        self.path_lbl = _ml_label(pr, "(데이터 서버 루트)", F_BASE, False, ML["muted"])
        self.path_lbl.pack(side="left")
        self.broom_btn = IconBtn(pr, "broom", self.clear_all, bg=ML["panel"], size=26)
        self.broom_btn.pack(side="right")
        _ml_line(inner).pack(fill="x")

        # 아래 줄: 개수 + 갱신 시각
        _ml_line(inner).pack(side="bottom", fill="x")
        ft = tk.Frame(inner, bg=ML["panel"])
        ft.pack(side="bottom", fill="x", padx=8, pady=5)
        self.status_lbl = _ml_label(ft, "꺼짐", F_BASE, False, ML["muted"])      # 접속 상태: IP 앞
        self.status_lbl.pack(side="left", padx=(0, 10))
        self.meta_lbl = _ml_label(ft, site["host"], F_BASE, False, ML["dim"])
        self.meta_lbl.pack(side="left")
        self.count_lbl = _ml_label(ft, "", F_BASE, False, ML["dim"])
        self.count_lbl.pack(side="left", padx=(12, 0))
        self.stamp_lbl = _ml_label(ft, "", F_SMALL, False, ML["dim"])
        self.stamp_lbl.pack(side="right")

        wrap, self.tree, self.empty = _file_list(inner, self.enter, PC_COLS)
        wrap.pack(fill="both", expand=True)
        self.list_wrap = wrap
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self._update_count())
        self.tree.bind("<Button-3>", self._open_text)
        tw.attach(self, self.tree, self.card)

    # ---- 선택 / 끌어놓기 공통
    def has_selection(self):
        return bool(self.tree.selection())

    def clear_selection(self):
        if self.tree.selection():
            self.tree.selection_set(())

    def selected(self):
        return [self.rows[i] for i in self.tree.selection() if i in self.rows]

    def drag_files(self):
        return [e for e in self.selected() if not e.is_dir]

    # ---- 연결
    def connect(self):
        if self.session is not None:
            return
        self.session = FtpSession(self.key, self.site["host"], self.site["user"], self.site["password"],
                                  self.tw.events, retry_s=self.retry_s, start_cwd=self.cwd)
        self.session.start()

    def disconnect(self):
        if self.session is not None:
            self.session.stop()
            self.session = None
        self._state("off", "")

    def _send(self, *cmd):
        if self.session is not None and self.state == "on":
            self.session.submit(*cmd)
            return True
        self.tw.say(f"{self.name}: 접속돼 있지 않습니다")
        return False

    # ---- 이벤트
    def on_event(self, kind, payload):
        if kind == "state":
            self._state(payload[0], payload[1], payload[2])
        elif kind == "list":
            cwd, entries = payload
            sig = (cwd, tuple((e.name, e.is_dir, e.size, e.mtime) for e in entries))
            self.cwd, self.entries = cwd, entries
            self.tw.cwd_mem[self.name] = self.cwd
            if sig != getattr(self, "_list_sig", None):          # 달라졌을 때만 다시 그림 (깜박임·선택 풀림 방지)
                self._list_sig = sig
                self._render()
            else:
                self.stamp_lbl.config(text="마지막 갱신 " + time.strftime("%H:%M:%S"))
            self.busy = False
        elif kind == "progress":
            label, done, total, idx, count = payload
            pct = int(done * 100 / total) if total else 0
            self.busy = True
            self.count_lbl.config(text=f"{label}  {pct}%  ({idx}/{count})")
            self.tw.progress(pct, f"{self.name} · {label}  {pct}%  ({idx}/{count})")
        elif kind == "done":
            op, ok, bad = payload
            _po = self.tw.pending_open
            _view = bool(_po and op == "download" and _po["panel"] is self)      # 오른쪽 클릭 미리보기용 내려받기
            verb = {"upload": "올리기", "download": "내려받기", "delete": "삭제"}.get(op, op)
            self.tw.log(self.name, f"{verb} {ok}개 완료" + (f", {bad}개 실패" if bad else ""))
            self.tw.progress(0, "준비")
            self.busy = False
            self._update_count()
            if op == "download":
                self.tw.pc.refresh()
            self.tw.relay_step(self, op)
            self.tw.open_step(self, op)
            if op in ("upload", "download") and ok > 0 and not bad and not _view:
                self.tw.auto_min_soon()
        elif kind == "error":
            self.busy = False
            self.tw.log(self.name, payload[1], error=True)
            self.tw.say(f"{self.name}: {payload[1]}")
            self.tw.progress(0, "준비")
            self._update_count()
        elif kind == "uploaded":
            if payload in getattr(self, "_del_after", ()):
                try:
                    os.remove(payload)
                    self.tw.log("내 PC", f"{Path(payload).name} 전송 후 삭제")
                except OSError as e:
                    self.tw.log("내 PC", f"{Path(payload).name}: 지우지 못했습니다 ({e.strerror or e})", error=True)
                self.tw.pc.refresh()
        elif kind == "log":
            self.tw.log(self.name, payload)

    def _state(self, st, msg="", ms=0):
        self.state = st
        if st == "on":
            self.status_lbl.config(text="접속됨")
            self.meta_lbl.config(text=f"{self.site['host']}  ·  {ms} ms")
        elif st == "connecting":
            self.status_lbl.config(text="접속 중…")
            self.meta_lbl.config(text=self.site["host"])
        else:
            deny = msg.startswith("로그인 실패")
            self.status_lbl.config(text=("로그인 실패" if deny else (f"꺼짐 · {self.retry_s}초 뒤 다시" if msg else "꺼짐")))
            self.meta_lbl.config(text=self.site["host"])
            if msg:
                self.tw.log(self.name, msg, error=True)
                self.entries = []
                self._render()

    # ---- 목록
    def _render(self):
        sel = {self.rows[i].name for i in self.tree.selection() if i in self.rows}
        top = self.tree.yview()[0]
        self.tree.delete(*self.tree.get_children())
        self.rows = {}
        ordered = sorted(self.entries, key=lambda e: (not e.is_dir, e.name.casefold()))
        for n, e in enumerate(ordered):
            iid = str(n)
            self.rows[iid] = e
            self.tree.insert("", "end", iid=iid, tags=self._tags(n, e), values=self._vals(e))
            if e.name in sel:
                self.tree.selection_add(iid)
        if ordered:
            self.empty.place_forget()
        else:
            self.empty.place(relx=0.5, rely=0.5, anchor="center")
        self.tree.yview_moveto(top)
        self.path_lbl.config(text="(데이터 서버 루트)" if self.cwd in ("/", "") else self.cwd)
        self.stamp_lbl.config(text="마지막 갱신 " + time.strftime("%H:%M:%S"))
        self._update_count()

    def _is_running(self, e):
        return False

    def _tags(self, n, e):
        return (("dir",) if e.is_dir else ()) + (("alt",) if n % 2 else ())

    def _vals(self, e):
        name = ("▸  " if e.is_dir else "") + e.name
        return (name, "" if e.is_dir else _fmt_size(e.size), "" if e.is_dir else _fmt_time(e.mtime))

    def _update_count(self):
        if not self.busy:
            self.count_lbl.config(text="")

    def _open_text(self, e):
        """오른쪽 클릭: 그 파일을 메모장으로 열어 본다 (장비 파일은 임시로 내려받아서, 읽기 전용)."""
        row = self.tree.identify_row(e.y)
        ent = self.rows.get(row)
        if ent is None:
            return
        if row not in self.tree.selection():
            self.tree.selection_set(row)
        if ent.is_dir:
            self.tw.say("폴더는 열 수 없습니다")
            return
        tmp = self.tw.new_view_dir()
        self.tw.pending_open = {"panel": self, "path": tmp / norm_prog_name(ent.name)}
        self.start_download([ent], tmp, ask=False)

    # ---- 동작
    def refresh(self):
        self._send("list")

    def auto_refresh(self):
        """보이는 동안 주기적으로 장비 목록을 다시 읽는다 (다른 PC 가 바꾼 파일도 곧 보임)."""
        if self.state == "on" and not self.busy and self.tw._edit is None and self.tw._drag is None \
                and self.session is not None:
            self.session.submit("list")

    def click_refresh(self):
        """목록을 누르면 새로고침 (1초에 한 번, 접속 중일 때만)."""
        now = time.monotonic()
        if self.state == "on" and not self.busy and now - self._last_click > 1.0 and self.tw._edit is None:
            self._last_click = now
            self.session.submit("list")

    def up(self):
        if self.cwd not in ("/", ""):
            self._send("cd", "..")

    def enter(self, iid):
        e = self.rows.get(iid)
        if e and e.is_dir:
            self._send("cd", e.name)

    def start_upload(self, paths, ask=True):
        """내 PC 파일들을 이 장비(현재 폴더)로 올린다."""
        paths = [Path(p) for p in paths if Path(p).is_file()]
        if not paths:
            self.tw.say("올릴 파일이 없습니다 (폴더는 올리지 않습니다)")
            return
        exist = {e.name.casefold() for e in self.entries}
        clash = [p for p in paths if norm_prog_name(p.name).casefold() in exist]
        if ask and clash and not _ask_overwrite(self.tw, "올리기", [p.name for p in clash]):
            paths = [p for p in paths if p not in clash]
        self._del_after = set()
        if paths and getattr(self.tw, "relay", None) is None and self.tw.pc.del_var.get():
            self._del_after = {str(p) for p in paths}          # '전송 시 삭제' 체크: 올라간 파일은 내 PC 에서 지움
        if paths and self._send("upload", paths):
            self.busy = True
            self.tw.log(self.name, f"{len(paths)}개 올리는 중…")

    def clear_all(self):
        """빗자루: 이 장비의 현재 폴더 파일을 모두 지운다 (폴더가 있으면 하지 않고, 실행 중인 프로그램은 남긴다)."""
        if self.state != "on" or self.busy:
            self.tw.say(f"{self.name}: 지금은 지울 수 없습니다 (접속 중이 아니거나 작업 중)")
            return
        if any(e.is_dir for e in self.entries):
            self.tw.say(f"{self.name}: 폴더가 있어서 전체 삭제를 하지 않습니다")
            return
        files = [e for e in self.entries if not self._is_running(e)]
        kept = len(self.entries) - len(files)
        if not files:
            self.tw.say(f"{self.name}: 지울 파일이 없습니다" + (" (실행 중인 프로그램만 있음)" if kept else ""))
            return
        self._send("delete", [e.name for e in files])           # 확인 없이 바로 실행

    def start_download(self, files, dest, ask=True):
        """장비의 파일들을 내 PC 폴더 dest 로 내려받는다."""
        dest = Path(dest)
        files = [e for e in files if not e.is_dir]
        if not files:
            self.tw.say("내려받을 파일이 없습니다 (폴더는 내려받지 않습니다)")
            return
        clash = [e for e in files if (dest / norm_prog_name(e.name)).exists()]
        if ask and clash and not _ask_overwrite(self.tw, "내려받기", [e.name for e in clash]):
            files = [e for e in files if e not in clash]
        if files and self._send("download", [e.name for e in files], dest):
            self.busy = True
            self.tw.log(self.name, f"{len(files)}개 내려받는 중…")

    def delete(self):
        files = self.selected()
        if not files:
            self.tw.say("지울 파일을 선택하세요")
            return
        shown = ", ".join(e.name for e in files[:6]) + (" …" if len(files) > 6 else "")
        dirs = sum(1 for e in files if e.is_dir)
        note = f"\n(폴더 {dirs}개는 안의 파일까지 모두 지워집니다)" if dirs else ""
        if messagebox.askyesno("삭제", f"장비 '{self.name}' 에서 {len(files)}개를 지웁니다. 되돌릴 수 없습니다.{note}\n{shown}\n\n계속할까요?",
                               icon="warning", parent=self.tw):
            self._send("delete", [e.name for e in files])

    def rename(self):
        files = self.selected()
        if len(files) != 1:
            self.tw.say("이름을 바꿀 파일을 하나만 선택하세요")
            return
        old = files[0].name
        self.tw.inline_rename(self.tree, self.tree.selection()[0], old, lambda new: self._send("rename", old, new))


# ---------------------------------------------------------------- 내 PC 패널 (아래쪽)
def _is_hidden(entry):
    try:
        return bool(getattr(entry.stat(follow_symlinks=False), "st_file_attributes", 0) & 0x6)
    except OSError:
        return False


def pc_drives():
    if IS_WINDOWS:
        mask = ctypes.windll.kernel32.GetLogicalDrives()
        return [f"{chr(65 + i)}:\\" for i in range(26) if mask >> i & 1]
    return ["/"]


def pc_places(cfg):
    home = Path.home()
    out = [("NCDATA", ncdata_dir(cfg))]
    od = os.environ.get("OneDrive")
    for label, names in (("바탕 화면", ("Desktop", "바탕 화면")), ("다운로드", ("Downloads", "다운로드")),
                         ("문서", ("Documents", "문서")), ("사진", ("Pictures", "사진"))):
        base = _documents_dir() if label == "문서" else None
        cands = [base] if base else []
        for root in (home, Path(od) if od else None):
            if root:
                cands += [root / n for n in names]
        path = next((c for c in cands if c and c.is_dir()), None)
        if path:
            out.append((label, path))
    return out


class PcPanel:
    title = "내 PC"

    def __init__(self, inner, tw, start):
        self.tw = tw
        self.root = inner
        self.base = Path(start)
        self.cwd = Path(start)
        self.rows = {}
        self._mtime = None
        self._last_click = 0.0

        hd = tk.Frame(inner, bg=ML["panel"])
        hd.pack(fill="x", padx=8, pady=6)
        self.title_lbl = _ml_label(hd, "내 PC", F_NAME, True)
        self.title_lbl.pack(side="left")
        self.path_lbl = _ml_label(hd, str(self.cwd), F_BASE, False, ML["muted"])      # 화면에는 표시하지 않음
        self.target = 0                                  # 숫자키로 고른 전송 대상 장비 번호 (0 = 없음)
        self.target_lbl = _ml_label(hd, "", F_NAME, True, ML["blue"])
        self.target_lbl.pack(side="left", padx=(10, 0))
        self.broom_btn = IconBtn(hd, "broom", self.clear_all, bg=ML["panel"], size=26)
        self.broom_btn.pack(side="right")
        self.del_var = tk.BooleanVar(value=bool(tw.cfg().get("del_after_send", False)))
        self.del_pill = PillToggle(hd, "전송 후 삭제", self.del_var, self._save_del)
        self.del_pill.pack(side="right", padx=(0, 8))
        _ml_line(inner).pack(fill="x")
        _ml_line(inner).pack(side="bottom", fill="x")
        ft = tk.Frame(inner, bg=ML["panel"])
        ft.pack(side="bottom", fill="x", padx=8, pady=5)
        self.count_lbl = _ml_label(ft, "", F_BASE, False, ML["dim"])
        self.count_lbl.pack(side="left")
        self.stamp_lbl = _ml_label(ft, "", F_SMALL, False, ML["dim"])
        self.stamp_lbl.pack(side="right")
        wrap, self.list, self.empty = _file_list(inner, self._open_row, PC_COLS)
        wrap.pack(fill="both", expand=True)
        self.list_wrap = wrap
        self.list.bind("<Return>", self._enter)
        self.list.bind("<KeyPress>", self._digit, add="+")
        self.list.bind("<<TreeviewSelect>>", lambda _e: self._update_count())
        self.list.bind("<Button-3>", self._open_text)
        self.list.bind("<BackSpace>", lambda _e: self.go_up())          # 상위 폴더 (Backspace)
        tw.attach(self, self.list, inner)
        self.navigate(self.cwd)

    # ---- 선택 / 끌어놓기 공통
    def has_selection(self):
        return bool(self.list.selection())

    def clear_selection(self):
        if self.list.selection():
            self.list.selection_set(())

    def drag_files(self):
        return self.selected_files()

    def selected_files(self):
        return [self.cwd / self.rows[i][0] for i in self.list.selection()
                if i in self.rows and not self.rows[i][1]]

    def drop_dir(self, x_root, y_root):
        """끌어다 놓은 위치가 폴더 줄 위면 그 폴더, 아니면 지금 보는 폴더."""
        try:
            y = y_root - self.list.winfo_rooty()
            x = x_root - self.list.winfo_rootx()
            if 0 <= x < self.list.winfo_width() and 0 <= y < self.list.winfo_height():
                r = self.rows.get(self.list.identify_row(y))
                if r and r[1]:
                    return self.cwd.parent if r[0] == ".." else self.cwd / r[0]
        except tk.TclError:
            pass
        return self.cwd

    def click_refresh(self):
        """목록을 누르면 새로고침 (너무 자주 하지 않도록 1초에 한 번)."""
        now = time.monotonic()
        if now - self._last_click > 1.0 and self.tw._edit is None:
            self._last_click = now
            self.refresh()

    # ---- 이동 / 목록
    def navigate(self, path):
        path = Path(path)
        if not path.is_dir():
            self.tw.say(f"폴더를 열 수 없습니다: {path}")
            return
        self.cwd = path
        cfg = self.tw.cfg()
        if cfg.get("pc_last") != str(path):                # 마지막으로 본 폴더를 기억 (다음 실행 때 거기서 시작)
            cfg["pc_last"] = str(path)
            save_app_cfg(cfg)
        self.path_lbl.config(text=str(path))
        self.list.selection_set(())
        self.refresh()
        self.list.yview_moveto(0)

    def go_up(self):
        if self.cwd.parent != self.cwd:
            self.navigate(self.cwd.parent)

    def refresh(self, select=None):
        from datetime import datetime
        sel = {self.rows[i][0] for i in self.list.selection() if i in self.rows}
        if select:                                       # 이름을 바꾼 파일을 새 이름으로 계속 선택
            sel = {select}
        top = self.list.yview()[0]
        self.list.delete(*self.list.get_children())
        self.rows = {}
        items = []
        try:
            for x in os.scandir(self.cwd):
                if _is_hidden(x):
                    continue
                try:
                    st = x.stat()
                    isdir = x.is_dir()
                except OSError:
                    continue
                items.append((not isdir, x.name.casefold(), x.name, isdir, st.st_size, datetime.fromtimestamp(st.st_mtime)))
        except OSError as e:
            self.tw.say(f"폴더를 읽지 못했습니다: {e.strerror or e}")
        items.sort()
        n = 0
        for _, _, name, isdir, size, mtime in items:
            iid = str(n)
            self.rows[iid] = (name, isdir)
            self.list.insert("", "end", iid=iid, tags=(("dir",) if isdir else ()) + (("alt",) if n % 2 else ()),
                             values=(("▸  " if isdir else "") + name, "" if isdir else _fmt_size(size),
                                     "" if isdir else _fmt_time(mtime)))
            if name in sel:
                self.list.selection_add(iid)
                if select:
                    self.list.focus(iid)
                    keep_see = iid
            n += 1
        if self.rows:
            self.empty.place_forget()
        else:
            self.empty.place(relx=0.5, rely=0.5, anchor="center")
        self.list.yview_moveto(top)
        if select and "keep_see" in locals():
            self.list.see(keep_see)
        try:
            self._mtime = self.cwd.stat().st_mtime
        except OSError:
            self._mtime = None
        self.stamp_lbl.config(text="마지막 갱신 " + time.strftime("%H:%M:%S"))
        self._update_count()

    def poll(self):
        """폴더 안 파일이 바뀌었으면 (다른 프로그램·다른 PC 가 넣거나 고쳤을 때) 목록을 다시 그린다."""
        try:
            with os.scandir(self.cwd) as it:
                sig = tuple(sorted((x.name, x.stat().st_mtime_ns, x.stat().st_size if x.is_file() else 0) for x in it))
        except OSError:
            return
        if sig != getattr(self, "_psig", None):
            self._psig = sig
            self.refresh()

    def _update_count(self):
        self.count_lbl.config(text="")

    def _open_text(self, e):
        row = self.list.identify_row(e.y)
        r = self.rows.get(row)
        if r is None:
            return
        if row not in self.list.selection():
            self.list.selection_set(row)
        if r[1]:
            self.tw.say("폴더는 열 수 없습니다")
        elif not open_in_notepad(self.cwd / r[0]):
            self.tw.say("프로그램을 열지 못했습니다")

    def _open_row(self, iid):
        r = self.rows.get(iid)
        if r and r[1]:
            self.navigate(self.cwd.parent if r[0] == ".." else self.cwd / r[0])

    # ---- 숫자키 + 엔터: 고른 번호의 장비(메인창 왼쪽부터 1, 2, 3…)로 파일 보내기
    def _devices(self):
        return [p for _k, p in sorted(self.tw.panels.items())]

    def _save_del(self):
        cfg = self.tw.cfg()
        cfg["del_after_send"] = bool(self.del_var.get())
        save_app_cfg(cfg)

    def clear_all(self):
        """빗자루: 내 PC 의 현재 폴더(D:\\NCDATA) 안의 파일을 모두 지운다 (폴더가 있으면 하지 않음)."""
        try:
            items = list(self.cwd.iterdir())
        except OSError as e:
            self.tw.say(f"내 PC: 폴더를 읽지 못했습니다 ({e.strerror or e})")
            return
        if any(i.is_dir() for i in items):
            self.tw.say("내 PC: 폴더가 있어서 전체 삭제를 하지 않습니다")
            return
        ok = 0
        for f in items:
            try:
                os.remove(f)
                ok += 1
            except OSError as e:
                self.tw.log("내 PC", f"{f.name}: 지우지 못했습니다 ({e.strerror or e})", error=True)
        self.tw.log("내 PC", f"전체 삭제 {ok}개 완료")
        self.refresh()

    def _show_target(self):
        devs = self._devices()
        if len(devs) == 1:                               # 장비가 하나뿐이면 번호를 안 눌러도 그 장비로 보냄
            self.target = 1
        if not self.target or self.target > len(devs):
            self.target = 0
            self.target_lbl.config(text="")
        else:
            self.target_lbl.config(text=devs[self.target - 1].name)

    def _digit(self, e):
        ch = e.char
        if not ch or not ch.isdigit() or e.state & 0x4:
            return
        n = int(ch)
        if n == 0:
            self.target = 0                              # 0 = 대상 해제
        elif n <= len(self._devices()):
            self.target = n                              # 장비 수보다 큰 숫자는 무시
        else:
            return "break"
        self._show_target()
        return "break"

    def _enter(self, _e=None):
        self._show_target()
        files = [r[0] for r in self._selected_rows() if not r[1]]
        if not self.target or not files:
            self._open_row(self.list.focus())            # 번호가 없으면 예전처럼 폴더 열기
            return "break"
        dev = self._devices()[self.target - 1]
        dev.start_upload([self.cwd / n for n in files])
        return "break"

    def _selected_rows(self):
        return [self.rows[i] for i in self.list.selection() if i in self.rows and self.rows[i][0] != ".."]

    # ---- 동작 (F2 / Del)
    def _selected_rows(self):
        return [self.rows[i] for i in self.list.selection() if i in self.rows]

    def rename(self):
        rows = self._selected_rows()
        if len(rows) != 1:
            self.tw.say("이름을 바꿀 파일을 하나만 선택하세요")
            return
        old = rows[0][0]
        self.tw.inline_rename(self.list, self.list.selection()[0], old, lambda new: self._rename_to(old, new))

    def _rename_to(self, old, new):
        try:
            if (self.cwd / new).exists():
                raise OSError("같은 이름이 이미 있습니다")
            os.rename(self.cwd / old, self.cwd / new)
            self.tw.log("내 PC", f"{old} → {new} 이름 바꿈")
        except OSError as e:
            self.tw.log("내 PC", f"이름을 바꾸지 못했습니다: {e.strerror or e}", error=True)
            new = old
        self.refresh(select=new)
        self.list.focus_set()                            # 바로 Enter 로 전송할 수 있게

    def delete(self):
        rows = self._selected_rows()
        if not rows:
            self.tw.say("지울 파일을 선택하세요")
            return
        files = [r[0] for r in rows if not r[1]]
        skipped = len(rows) - len(files)
        if not files:
            self.tw.say("폴더는 여기서 지우지 않습니다 (파일만 지워집니다)")
            return
        shown = ", ".join(files[:6]) + (" …" if len(files) > 6 else "")
        note = f"\n(폴더 {skipped}개는 지우지 않습니다)" if skipped else ""
        if not messagebox.askyesno("삭제", f"내 PC 에서 {len(files)}개를 지웁니다. 되돌릴 수 없습니다.{note}\n{shown}\n\n계속할까요?",
                                   icon="warning", parent=self.tw):
            return
        ok = 0
        for name in files:
            try:
                os.remove(self.cwd / name)
                ok += 1
            except OSError as e:
                self.tw.log("내 PC", f"{name}: 지우지 못했습니다 ({e.strerror or e})", error=True)
        self.tw.log("내 PC", f"삭제 {ok}개 완료")
        self.refresh()

    def copy_in(self, paths, dest):
        """다른 곳에서 끌어다 놓은 파일을 dest 폴더로 복사."""
        dest = Path(dest)
        paths = [Path(p) for p in paths if Path(p).is_file() and Path(p).parent != dest]
        if not paths:
            return
        clash = [p for p in paths if (dest / p.name).exists()]
        if clash and not _ask_overwrite(self.tw, "복사", [p.name for p in clash]):
            paths = [p for p in paths if p not in clash]
        ok = 0
        for p in paths:
            try:
                shutil.copy2(p, dest / p.name)
                ok += 1
            except OSError as e:
                self.tw.log("내 PC", f"{p.name}: 복사하지 못했습니다 ({e.strerror or e})", error=True)
        self.tw.log("내 PC", f"복사 {ok}개 완료 → {dest.name or dest}")
        self.refresh()


# ---------------------------------------------------------------- 파일 전송 창
class TransferWindow(tk.Toplevel):
    MIN_W, MIN_H = 900, 560
    REFRESH_SHOWN_SEC = 3.0         # 보이는 동안 목록 새로고침 간격 (설정에서 바꿈)
    REFRESH_HIDDEN_SEC = 30.0       # 숨겨진 동안

    def __init__(self, monitor):
        super().__init__(monitor)
        self.monitor = monitor
        self.withdraw()
        self.overrideredirect(True)                 # 제목표시줄 없음 + 작업표시줄에 안 나옴
        self.configure(bg=ML["bg"])
        ml_styles(self)
        self.events = queue.Queue()
        self.panels = {}
        self._log_lines = 0
        self._dirty = True
        self._tick_id = None
        self._t_run = self._t_pc = 0.0
        self._roots = {}                  # 끌어놓기 대상 영역: 위젯 경로 -> 패널
        self._drag = None
        self._ghost = None
        self._ghost_lbl = None
        self._sel_lock = False
        self.active = None                # 마지막으로 선택한 목록의 패널 (F2 / Del 대상)
        self.relay = None                 # 장비 -> 장비 복사 중인 작업
        self.cwd_mem = {}                 # 장비별로 마지막에 보던 폴더
        self._view_dirs = []              # 메모장으로 열어 본 임시 폴더들
        self.pending_open = None
        self._edit = None                 # 이름 바꾸는 중인 입력칸
        self._full = False
        self._prev_geo = None
        self._appwin = False
        self.title("FTP Lite")
        g = monitor.app_cfg.get("tw_geom")
        try:
            x, y, w, h = (int(v) for v in g)
            ok = on_screen(x, y)
        except (TypeError, ValueError):
            ok = False
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        if not ok:
            w, h = min(1180, sw - 80), min(780, sh - 120)
            x, y = max(0, (sw - w) // 2), max(0, (sh - h) // 3)
        self.geo = [x, y, max(self.MIN_W, w), max(self.MIN_H, h)]
        self._apply_geo()

        # 제목줄 (끌어서 이동)
        tb = tk.Frame(self, bg=ML["panel3"], height=44)
        tb.pack(fill="x")
        tb.pack_propagate(False)
        t1 = _ml_label(tb, "FTP Lite", F_NAME, True, ML["brand"], ML["panel3"])
        t1.pack(side="left", padx=(12, 6))
        IconBtn(tb, "close", self.hide, bg=ML["panel3"]).pack(side="right", padx=(0, 10))
        self.full_btn = IconBtn(tb, "full", self.toggle_full, bg=ML["panel3"])
        self.full_btn.pack(side="right")
        IconBtn(tb, "sun" if DARK_MODE[0] else "moon", lambda: self.monitor.set_dark(not DARK_MODE[0]),
                bg=ML["panel3"]).pack(side="right")
        IconBtn(tb, "gear", self.open_settings, bg=ML["panel3"]).pack(side="right", padx=(0, 10))
        for w_ in (tb, t1):
            w_.bind("<ButtonPress-1>", self._move_start)
            w_.bind("<B1-Motion>", self._move)
            w_.bind("<ButtonRelease-1>", lambda _e: self._save_geo())

        # 상태줄
        sb = tk.Frame(self, bg=ML["panel3"], height=26)
        sb.pack(side="bottom", fill="x")
        sb.pack_propagate(False)
        self.status = _ml_label(sb, "준비", F_BASE, False, ML["muted"], ML["panel3"])
        self.status.pack(side="left", padx=10)
        grip = tk.Label(sb, text="◢", font=ui_font(10), bg=ML["panel3"], fg=ML["dim"], cursor="size_nw_se" if IS_WINDOWS else "bottom_right_corner")
        grip.pack(side="right", padx=(4, 4))
        grip.bind("<ButtonPress-1>", self._size_start)
        grip.bind("<B1-Motion>", self._size)
        grip.bind("<ButtonRelease-1>", lambda _e: self._save_geo())
        self.bar = ttk.Progressbar(sb, length=180, mode="determinate", value=0, style="ML.Horizontal.TProgressbar")
        self.bar.pack(side="right", padx=6)

        body = tk.Frame(self, bg=ML["bg"])
        body.pack(fill="both", expand=True, padx=8, pady=(5, 6))
        self.split = tk.PanedWindow(body, orient="vertical", bg=ML["bg"], bd=0, sashwidth=6, sashrelief="flat",
                                    opaqueresize=True)
        self.split.pack(fill="both", expand=True)
        self.top = tk.Frame(self.split, bg=ML["bg"])
        self.top.rowconfigure(0, weight=1)
        mid = tk.Frame(body, bg=ML["bg"])
        mid.pack(side="bottom", fill="x", before=self.split)       # 로그창: 3줄 높이로 고정 (끌어서 못 바꿈)
        self._build_log(mid)
        self._mid = mid
        bottom = tk.Frame(self.split, bg=ML["bg"])
        card = MPanel(bottom)
        card.pack(fill="both", expand=True, pady=(4, 0))
        self.pc = PcPanel(card.inner, self, self._pc_initial())
        self.pc.base = self._pc_start()
        self.split.add(self.top, minsize=180, stretch="always")
        self.split.add(bottom, minsize=160, stretch="always")
        # 설비 칸과 내 PC(히스토리) 칸은 세로로 정확히 반반 고정 (끌어서 못 바꿈, 전송 기록이 숨어도 그대로)
        for ev in ("<Button-1>", "<B1-Motion>", "<ButtonRelease-1>"):
            self.split.bind(ev, lambda _e: "break")
        self._half_job = None
        self.split.bind("<Configure>", self._half_soon)
        self.bind_all("<ButtonPress-1>", self._ensure_focus, add="+")
        self.bind("<F2>", lambda _e: self._key("rename"))
        self.bind("<Delete>", lambda _e: self._key("delete"))
        self.bind("<F5>", lambda _e: self._key("refresh"))
        self.bind("<Control-a>", self._select_all)
        self.bind("<Control-A>", self._select_all)
        self.protocol("WM_DELETE_WINDOW", self.monitor.quit_app)          # Alt+F4 / 닫기 요청 = 프로그램 종료
        self.bind("<Alt-F4>", lambda _e: (self.monitor.quit_app(), "break")[1])
        self.bind("<Escape>", lambda _e: self.hide())              # ESC: 큰 화면에서도 바로 최소화(숨김)
        self._tick()

    def refresh_shown(self):
        return cfg_seconds(self.monitor.app_cfg, "refresh_shown_sec", 1, low=1, high=3600)

    def refresh_hidden(self):
        return cfg_seconds(self.monitor.app_cfg, "refresh_hidden_sec", 30, low=1, high=3600)

    # ---- 창 이동 / 크기
    def _apply_geo(self):
        x, y, w, h = self.geo
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.after(30, self._round)

    CORNER_R = 8                                    # 창 모드일 때 모서리 둥글기(px)

    def _round(self):
        round_window(self, self.CORNER_R, self._full)

    def open_settings(self):
        self.monitor.open_settings()
        w = self.monitor.settings_win
        if w is not None and w.winfo_exists():
            w.lift()

    def _half_soon(self, _e=None):
        if self._half_job is None:
            self._half_job = self.after_idle(self._half)

    def _half(self):
        self._half_job = None
        try:
            h = self.split.winfo_height()
            if h > 50:
                self.split.sash_place(0, 0, h // 2)
        except tk.TclError:
            pass

    def sync_topmost(self):
        """메인 창은 '항상 위'를 쓰지 않는다 (일반 창). 열 때만 앞으로 올린다."""
        try:
            self.attributes("-topmost", False)
        except tk.TclError:
            pass

    def toggle_full(self):
        if self._full:
            self._full = False
            if self._prev_geo:
                self.geo = list(self._prev_geo)
            self.full_btn.set_icon("full")
        else:
            self._prev_geo = list(self.geo)
            self._full = True
            self.geo = list(work_area(self))
            self.full_btn.set_icon("restore")
        self._apply_geo()

    def _taskbar_entry(self):
        """제목표시줄 없는 창도 작업표시줄(시작표시줄)에 나오게 한다 (Windows)."""
        if not IS_WINDOWS or self._appwin:
            return
        self._appwin = True
        try:
            u = ctypes.WinDLL("user32", use_last_error=True)
            u.GetParent.argtypes, u.GetParent.restype = [wintypes.HWND], wintypes.HWND
            u.GetWindowLongW.argtypes, u.GetWindowLongW.restype = [wintypes.HWND, ctypes.c_int], ctypes.c_long
            u.SetWindowLongW.argtypes, u.SetWindowLongW.restype = [wintypes.HWND, ctypes.c_int, ctypes.c_long], ctypes.c_long
            hwnd = u.GetParent(self.winfo_id()) or self.winfo_id()
            ex = u.GetWindowLongW(hwnd, -20)
            u.SetWindowLongW(hwnd, -20, (ex | 0x00040000) & ~0x00000080)      # APPWINDOW 켜고 TOOLWINDOW 끔
            u.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
            u.SendMessageW.restype = ctypes.c_ssize_t
            self._hicons = (_make_icon(16), _make_icon(32))                 # 작업표시줄 버튼 아이콘: FTP
            u.SendMessageW(hwnd, 0x80, 0, self._hicons[0])
            u.SendMessageW(hwnd, 0x80, 1, self._hicons[1])
            try:                                                            # 창 클래스 아이콘도 같이 (작업표시줄이 클래스 아이콘을 쓰는 경우)
                u.SetClassLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
                u.SetClassLongPtrW.restype = ctypes.c_void_p
                u.SetClassLongPtrW(hwnd, -14, self._hicons[1])
                u.SetClassLongPtrW(hwnd, -34, self._hicons[0])
            except Exception:
                pass
            self.withdraw()                                                  # 다시 띄워야 작업표시줄에 등록된다
            self.update_idletasks()
            self.deiconify()
        except Exception:
            pass

    def _move_start(self, e):
        if self._full:
            return
        self._off = (e.x_root - self.geo[0], e.y_root - self.geo[1])

    def _move(self, e):
        if self._full:
            return
        self.geo[0], self.geo[1] = e.x_root - self._off[0], e.y_root - self._off[1]
        self._apply_geo()

    def _size_start(self, e):
        if self._full:
            return
        self._sz0 = (e.x_root, e.y_root, self.geo[2], self.geo[3])

    def _size(self, e):
        if self._full:
            return
        x0, y0, w0, h0 = self._sz0
        self.geo[2] = max(self.MIN_W, w0 + e.x_root - x0)
        self.geo[3] = max(self.MIN_H, h0 + e.y_root - y0)
        self._apply_geo()

    def _save_geo(self):
        if self._full:
            return
        self.monitor.app_cfg["tw_geom"] = list(self.geo)
        self.monitor._save_cfg()

    def _ensure_focus(self, _e=None):
        try:
            if self.winfo_viewable() and self.focus_displayof() is None:
                self.focus_force()
        except tk.TclError:                       # 창을 새로 만든 뒤 옛 창에 남은 연결
            pass

    # ---- 전송 기록 (가운데)
    def _build_log(self, parent):
        card = MPanel(parent)
        ls = tkfont.Font(font=ui_font(F_BASE)).metrics("linespace")
        card.configure(height=3 * (ls + 1) + 8 + 6)                  # 정확히 3줄 높이로 고정
        card.pack(fill="x", pady=(4, 0))
        self.logbox = tk.Text(card.inner, height=3, wrap="none", relief="flat", bd=0, bg=ML["panel"], fg=ML["text"],
                              font=ui_font(F_BASE), padx=10, pady=4, highlightthickness=0, state="disabled",
                              cursor="arrow", spacing1=1)
        self.logbox.pack(fill="both", expand=True)
        self.logbox.tag_configure("err", foreground=ML["red"])
        self.logbox.tag_configure("time", foreground=ML["dim"])

    def show_log(self, on):
        if on:
            self._mid.pack(side="bottom", fill="x", before=self.split)
        else:
            self._mid.pack_forget()
        self._half_soon()

    def log(self, name, text, error=False):
        self.logbox.config(state="normal")
        if self._log_lines:                      # 줄바꿈은 앞에 둔다 (끝에 두면 맨 아래에 빈 줄이 하나 생겨 2줄만 보임)
            self.logbox.insert("end", "\n")
        self.logbox.insert("end", time.strftime("%H:%M:%S") + "  ", "time")
        self.logbox.insert("end", f"{name} · {text}", "err" if error else ())
        self._log_lines += 1
        if self._log_lines > 300:
            self.logbox.delete("1.0", "2.0")
            self._log_lines -= 1
        self.logbox.see("end")
        self.logbox.config(state="disabled")

    def say(self, text):
        self.status.config(text=text)
        self.after(4000, lambda: self.status.config(text="준비") if self.status.cget("text") == text else None)

    def progress(self, pct, text):
        self.bar.configure(value=pct)
        self.status.config(text=text)

    # ---- 설정 / 설비
    def cfg(self):
        return self.monitor.app_cfg

    def _pc_initial(self):
        """창을 처음 열 때 시작할 폴더: 마지막으로 보던 폴더, 없어졌으면 기본 폴더."""
        last = self.cfg().get("pc_last")
        try:
            if last and Path(last).is_dir():
                return Path(last)
        except OSError:
            pass
        return self._pc_start()

    def _pc_start(self):
        p = pc_default_dir(self.cfg())
        return p if p.is_dir() else Path.home()

    def on_settings(self):
        """설정을 적용하면: 설비 목록은 다시 그려야 하므로 표시해 두고, 보이는 중이면 바로 다시 만든다."""
        self._dirty = True
        start = self._pc_start()
        if start != self.pc.base:                      # 설정에서 내 PC 기본 폴더를 바꿨으면 그리로
            self.pc.base = start
            self.pc.navigate(start)
        if self.winfo_viewable():
            self.rebuild()
            for p in self.panels.values():
                p.connect()

    # ---- 목록 선택 / 끌어놓기 / 키 처리
    def all_owners(self):
        pc = getattr(self, "pc", None)
        return [*self.panels.values(), *([pc] if pc is not None else [])]

    def attach(self, owner, tree, root):
        self._roots[str(root)] = owner
        owner.view = tree
        tree.bind("<ButtonPress-1>", lambda e: self._press(e, tree, owner))
        tree.bind("<B1-Motion>", lambda e: self._motion(e, owner))
        tree.bind("<ButtonRelease-1>", lambda e: self._release(e, owner))
        tree.bind("<<TreeviewSelect>>", lambda _e: self._selected(owner), add="+")

    def _selected(self, owner):
        """한 목록에서 고르면 다른 목록의 선택은 모두 푼다."""
        if self._sel_lock or not owner.has_selection():
            return
        self.active = owner
        self._sel_lock = True
        try:
            for o in self.all_owners():
                if o is not owner:
                    o.clear_selection()
        finally:
            self._sel_lock = False

    def _key(self, action):
        o, f = self.active, self.focus_get()
        if o is None or not isinstance(f, ttk.Treeview) or not o.has_selection() and action != "refresh":
            return
        if action == "refresh":
            (o.refresh if hasattr(o, "refresh") else (lambda: None))()
        else:
            getattr(o, action)()

    def _select_all(self, _e=None):
        f = self.focus_get()
        if isinstance(f, (tk.Entry, ttk.Entry, ttk.Combobox)):
            return
        o = next((x for x in self.all_owners() if getattr(x, "view", None) is f), None) or self.active
        if o is None or not hasattr(o, "view"):
            return "break"
        o.view.focus_set()
        o.view.selection_set(o.view.get_children())
        self.active = o
        return "break"

    def inline_rename(self, tree, iid, old, commit):
        """파일 탐색기처럼 그 자리에서 이름을 고친다 (Enter = 확인, Esc = 취소)."""
        self._end_edit()
        try:
            tree.see(iid)
            tree.update_idletasks()
            bx, by, bw, bh = tree.bbox(iid, "name")
        except (tk.TclError, ValueError):
            return
        var = tk.StringVar(value=old)
        ent = tk.Entry(tree, textvariable=var, font=ui_font(F_LIST), relief="solid", bd=1, highlightthickness=0,
                       bg=ML["panel"], fg=ML["text"], insertbackground=ML["text"])
        ent.place(x=bx, y=by, width=bw, height=bh)
        ent.focus_force()
        ent.select_range(0, "end")                      # 확장자까지 전체 선택
        ent.icursor("end")
        state = {"done": False}

        def finish(ok):
            if state["done"]:
                return
            state["done"] = True
            self._edit = None
            new = var.get().strip()
            try:
                ent.destroy()
            except tk.TclError:
                pass
            if ok and new and new != old:
                commit(new)
        ent.bind("<Return>", lambda _e: (finish(True), "break")[1])
        ent.bind("<Escape>", lambda _e: (finish(False), "break")[1])
        ent.bind("<FocusOut>", lambda _e: finish(True))
        self._edit = finish

    def _end_edit(self):
        if self._edit:
            self._edit(False)

    # ---- 오른쪽 클릭: 장비 파일을 임시로 내려받아 메모장으로
    def new_view_dir(self):
        d = Path(tempfile.mkdtemp(prefix="ncv_"))
        self._view_dirs.append(d)
        return d

    def open_step(self, panel, op):
        p = self.pending_open
        if p and op == "download" and p["panel"] is panel:
            self.pending_open = None
            if p["path"].is_file():
                try:
                    os.chmod(p["path"], 0o444)              # 읽기 전용 (고쳐도 장비에는 반영되지 않으므로)
                except OSError:
                    pass
                if not open_in_notepad(p["path"]):
                    self.say("프로그램을 열지 못했습니다")

    def _press(self, e, tree, owner):
        tree.focus_set()
        self._drag = None
        row = tree.identify_row(e.y)
        if tree.identify_region(e.x, e.y) in ("heading", "separator") or e.state & 0x5:
            return
        if not row or tree.identify_column(e.x) != "#1":
            # 빈 곳이나 이름 칸이 아닌 곳에서 시작한 끌기 = 여러 줄 선택 (이름 칸에서 끌면 파일 옮기기)
            kids = tree.get_children()
            self._drag = {"owner": owner, "tree": tree, "row": row, "x": e.x_root, "y": e.y_root, "on": False,
                          "defer": False, "band": True, "moved": False, "t": 0.0,
                          "ai": kids.index(row) if row in kids else len(kids)}
            return
        d = {"owner": owner, "tree": tree, "row": row, "x": e.x_root, "y": e.y_root, "on": False,
             "defer": row in tree.selection()}
        self._drag = d
        if d["defer"]:                       # 여러 개를 고른 채 끌 수 있도록, 놓을 때까지 선택을 바꾸지 않는다
            return "break"

    def _target_at(self, x, y):
        w = self.winfo_containing(x, y)
        while w is not None:
            o = self._roots.get(str(w))
            if o is not None:
                return o
            w = w.master
        return None

    def _band(self, d, e):
        """끌어서 여러 줄 선택 (위·아래 끝을 넘어가면 목록이 따라 움직임)."""
        tree = d["tree"]
        kids = tree.get_children()
        if not kids:
            return
        h = tree.winfo_height()
        y = e.y
        now = time.monotonic()
        if tree.identify_region(e.x, max(0, y)) == "heading" or y < 0:
            if now - d["t"] > 0.05:
                d["t"] = now
                tree.yview_scroll(-1, "units")
            y = 30
            while tree.identify_region(e.x, y) == "heading" and y < 80:
                y += 2
        elif y >= h:
            if now - d["t"] > 0.05:
                d["t"] = now
                tree.yview_scroll(1, "units")
            y = h - 1
        r = tree.identify_row(y)
        ci = kids.index(r) if r in kids else len(kids) - 1
        lo, hi = sorted((d["ai"], ci))
        hi = min(hi, len(kids) - 1)
        want = kids[lo:hi + 1]
        if tuple(tree.selection()) != tuple(want):
            tree.selection_set(want)

    def _motion(self, e, owner):
        d = self._drag
        if d is None:
            return
        if d.get("band"):
            if not d["moved"] and abs(e.x_root - d["x"]) + abs(e.y_root - d["y"]) < 5:
                return "break"
            d["moved"] = True
            self._band(d, e)
            return "break"
        if not d["on"]:
            if abs(e.x_root - d["x"]) + abs(e.y_root - d["y"]) < 8:
                return "break" if d["defer"] else None
            if not owner.drag_files():
                self._drag = None
                return
            d["on"] = True
            g = tk.Toplevel(self)
            g.overrideredirect(True)
            g.attributes("-topmost", True)
            self._ghost_lbl = tk.Label(g, bg=ML["blue"], fg=ML["on_accent"], font=ui_font(F_BASE, True), padx=10, pady=4)
            self._ghost_lbl.pack()
            self._ghost = g
        dst = self._target_at(e.x_root, e.y_root)
        n = len(owner.drag_files())
        self._ghost_lbl.config(text=f"{n}개 → {dst.title}" if dst is not None and dst is not owner
                               else (f"{n}개 → 같은 폴더" if dst is owner else f"{n}개"))
        self._ghost.geometry(f"+{e.x_root + 16}+{e.y_root + 12}")
        return "break"

    def _release(self, e, owner):
        d, self._drag = self._drag, None
        if d is None:
            owner.click_refresh()
            return
        if self._ghost is not None:
            self._ghost.destroy()
            self._ghost = None
        if d["on"]:
            dst = self._target_at(e.x_root, e.y_root)
            if dst is not None:
                self._drop(d["owner"], dst, e.x_root, e.y_root)
            return "break"
        if d.get("band") and d["moved"]:
            return "break"
        if d["defer"]:
            d["tree"].selection_set(d["row"])
        owner.click_refresh()
        if d["defer"]:
            return "break"

    def _drop(self, src, dst, x, y):
        files = src.drag_files()
        if not files:
            return
        if isinstance(src, MachinePanel):
            if isinstance(dst, PcPanel):
                src.start_download(files, dst.drop_dir(x, y))
            elif dst is not src:                                    # 장비 -> 장비: 임시 폴더를 거쳐 복사
                if dst.state != "on":
                    self.say(f"{dst.name}: 접속돼 있지 않습니다")
                    return
                self._relay_end()
                tmp = Path(tempfile.mkdtemp(prefix="ncm_"))
                self.relay = {"src": src, "dst": dst, "dir": tmp, "names": [norm_prog_name(f.name) for f in files]}
                src.start_download(files, tmp, ask=False)
        elif isinstance(dst, MachinePanel):
            dst.start_upload(files)
        elif isinstance(dst, PcPanel):
            dst.copy_in(files, dst.drop_dir(x, y))

    def relay_step(self, panel, op):
        r = self.relay
        if not r:
            return
        if op == "download" and r["src"] is panel and "phase" not in r:
            r["phase"] = "up"
            paths = [r["dir"] / n for n in r["names"] if (r["dir"] / n).is_file()]
            if paths and r["dst"].state == "on":
                r["dst"].start_upload(paths)
            else:
                self._relay_end()
        elif op == "upload" and r["dst"] is panel and r.get("phase") == "up":
            self._relay_end()

    def _relay_end(self):
        r, self.relay = self.relay, None
        if r:
            shutil.rmtree(r["dir"], ignore_errors=True)

    def rebuild(self):
        devices = [(n, s) for n, s in load_sites().items() if s.get("host") and s.get("ftp")]
        retry = cfg_seconds(self.cfg(), "reconnect_sec", DEFAULT_RECONNECT_SEC)
        old = {p.name: p for p in self.panels.values()}
        keep = {}
        for n, st in devices:                           # 이름·접속 정보가 같은 패널은 그대로 두고 깜빡임 없이 재사용
            op = old.get(n)
            if op is not None and (op.site.get("host"), op.site.get("user"), op.site.get("password")) == \
                    (st.get("host"), st.get("user"), st.get("password")):
                keep[n] = op
        for n, op in old.items():
            if keep.get(n) is not op:
                op.disconnect()
                op.card.destroy()
        pc = getattr(self, "pc", None)
        self._roots = {str(pc.root): pc} if pc is not None else {}
        self.active = None
        for w_ in self.top.winfo_children():
            if not any(w_ is op.card for op in keep.values()):
                w_.destroy()
        self.panels = {}
        for c in range(16):                              # 이전에 잡아 둔 칸 설정을 지운다 (남으면 한쪽만 차지)
            self.top.columnconfigure(c, weight=0, uniform="", minsize=0)
        if not devices:
            self.top.columnconfigure(0, weight=1, uniform="")
            msg = MPanel(self.top)
            msg.grid(row=0, column=0, sticky="nsew")
            _ml_label(msg.inner, "등록된 설비가 없습니다.\n설정 → 설비에서 장비를 추가하세요.",
                      F_BASE, False, ML["dim"], justify="center").pack(expand=True, pady=20)
            MButton(msg.inner, "설정 열기", self.monitor.open_settings).pack(pady=(0, 20))
        for i, (name, site) in enumerate(devices):
            self.top.columnconfigure(i, weight=1, uniform="m")            # 설비 패널 폭을 똑같이
            panel = keep.get(name)
            if panel is None:
                panel = MachinePanel(self.top, self, i, i + 1, name, site, retry)
            else:
                panel.key, panel.num, panel.site, panel.retry_s = i, i + 1, site, retry
                if panel.session is not None:
                    panel.session.key, panel.session.retry_s = i, retry
                self._roots[str(panel.card)] = panel
            panel.card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 4, 0))
            self.panels[i] = panel
        self._dirty = False
        if getattr(self, "pc", None) is not None:
            self.pc._show_target()

    # ---- 보이기 / 숨기기 (숨기면 장비 FTP 연결을 끊는다)
    def show(self):
        if self._dirty:
            self.rebuild()
        self.deiconify()
        self._apply_geo()
        self.update_idletasks()
        self._taskbar_entry()
        self.after(60, self._round)
        self.lift()
        self.after(150, self.sync_topmost)
        self.focus_force()
        self.pc.refresh()
        for p in self.panels.values():
            p.connect()

    def auto_min_soon(self):
        """전송이 끝나고 실패가 없으면 잠깐 뒤 창을 숨긴다 (설정에서 끌 수 있음)."""
        if bool(self.cfg().get("auto_min_done", True)):
            self.after(500, self._auto_min)

    def _auto_min(self):
        if not self.winfo_viewable() or self.relay is not None or self.pending_open is not None:
            return                                          # 장비 간 복사 중이거나 미리보기 대기 중이면 숨기지 않는다
        if any(p.busy for p in self.panels.values()):
            return                                          # 다른 전송이 이어지는 중
        self.hide()

    def hide(self):
        if not self.winfo_viewable():
            return
        self._end_edit()
        self._save_geo()
        for p in self.panels.values():
            p.disconnect()
        self.withdraw()

    def toggle(self):
        self.hide() if self.winfo_viewable() else self.show()

    def shutdown(self):
        self._relay_end()
        for d in self._view_dirs:
            for f in d.glob("*"):
                try:
                    os.chmod(f, 0o666)
                except OSError:
                    pass
            shutil.rmtree(d, ignore_errors=True)
        if self._tick_id:
            self.after_cancel(self._tick_id)
        for p in self.panels.values():
            p.disconnect()
        try:
            self.destroy()
        except tk.TclError:
            pass

    # ---- 주기 처리: 장비 이벤트 전달, 가동 상태 갱신, 내 PC 폴더 변화 확인
    def _tick(self):
        try:
            while True:
                key, kind, payload = self.events.get_nowait()
                p = self.panels.get(key)
                if p is not None:
                    p.on_event(kind, payload)
        except queue.Empty:
            pass
        now = time.monotonic()
        if self.winfo_viewable():
            if now - self._t_pc > self.refresh_shown():
                self._t_pc = now
                if self._edit is None:
                    self.pc.poll()
            if now - self._t_run > self.refresh_shown():
                self._t_run = now
                for p in self.panels.values():
                    p.auto_refresh()
        elif now - self._t_pc > self.refresh_hidden():       # 숨겨진 동안에는 30초마다 (내 PC 목록만. 장비 연결은 끊겨 있음)
            self._t_pc = now
            if self._edit is None:
                self.pc.poll()
        self._tick_id = self.after(150, self._tick)




class LiteApp(tk.Tk):
    """보이지 않는 바탕 창 + 트레이 아이콘 + 파일 전송 창 (모니터링 없음)."""

    def __init__(self):
        super().__init__()
        self.title("FTP_Lite")
        try:                                                  # 작업 표시줄 / 창 아이콘: 파란 'FTP'
            self._app_icons = [tk.PhotoImage(master=self, data=_ftp_icon_png(n)) for n in (64, 32, 16)]
            self.iconphoto(True, *self._app_icons)
        except tk.TclError:
            pass
        pick_ui_font(self)
        apply_styles(self)
        self.withdraw()
        self.app_cfg = load_app_cfg()
        apply_theme(bool(self.app_cfg.get("dark")))
        self.tw_hotkey = tw_hotkey_cfg(self.app_cfg)
        self.tw = None                      # 파일 전송 창
        self.settings_win = None
        self._first_run_pending = FIRST_RUN          # 설정 파일이 없는 첫 실행
        self._closing = False
        self.events = queue.Queue()
        self.tray = None
        if IS_WINDOWS:
            self.tray = TrayIcon(self.events, "FTP Lite", lambda: (False, False), self.tw_hotkey, None)
            if not self.tray.start():
                self.tray = None
        self._fallback_seq = None
        if self.tray is None:
            self._bind_fallback_hotkey()
        self.protocol("WM_DELETE_WINDOW", self.quit_app)
        self._poll_events()
        if self._first_run_pending:              # 처음 실행: 설정 + 장비 찾기를 먼저 연다
            self.after(600, self._first_run)
        else:
            self.after(300, self.show_transfer)      # 실행할 때마다 창을 보여 준다

    def report_callback_exception(self, exc, val, tb):
        """.pyw 는 콘솔이 없으므로, 예상 못 한 오류가 나면 창으로 알리고 파일에도 남긴다."""
        import traceback
        text = "".join(traceback.format_exception(exc, val, tb))
        try:
            (APP_DIR / "FTP_Lite_error.log").open("a", encoding="utf-8").write(time.strftime("%F %T\n") + text + "\n")
        except OSError:
            pass
        try:
            messagebox.showerror("FTP_Lite 오류", f"{val}\n\n자세한 내용: {APP_DIR / 'FTP_Lite_error.log'}")
        except tk.TclError:
            pass

    def _first_run(self):
        self.open_settings()
        if self.settings_win is not None and self.settings_win.winfo_exists():
            self.settings_win.after(400, self.settings_win.open_finder)

    def _bind_fallback_hotkey(self):
        if self._fallback_seq:
            self.unbind_all(self._fallback_seq)
        self._fallback_seq = hotkey_tk_sequence(self.tw_hotkey)
        if self._fallback_seq:
            self.bind_all(self._fallback_seq, lambda _e: self.toggle_transfer())

    def set_tw_hotkey(self, text):
        self.tw_hotkey = text
        if self.tray:
            self.tray.set_hotkey(text)
        else:
            self._bind_fallback_hotkey()

    def _save_cfg(self):
        save_app_cfg(self.app_cfg)

    def set_dark(self, dark):
        """다크/화이트 모드 전환 (버튼 이벤트가 끝난 뒤에 창을 새로 만든다)."""
        self.after(30, lambda: self._set_dark(dark))

    def _set_dark(self, dark):
        cfg = load_app_cfg()
        cfg["dark"] = bool(dark)
        save_app_cfg(cfg)
        apply_theme(dark)
        self.app_cfg = load_app_cfg()
        old = self.tw
        if old is None:
            return
        was_vis, was_full, mem = old.winfo_viewable(), old._full, dict(old.cwd_mem)
        if was_full:
            old.toggle_full()
        old._save_geo()
        old.shutdown()
        self.tw = None
        try:
            self.tw = TransferWindow(self)
        except Exception as e:
            messagebox.showerror("FTP Lite", f"파일 전송 창을 열지 못했습니다.\n{e}", parent=self)
            return
        self.tw.cwd_mem = mem
        if was_vis:
            self.tw.show()
            if was_full:
                self.tw.toggle_full()

    def toggle_transfer(self):
        if self.tw is None:
            try:
                self.tw = TransferWindow(self)
            except Exception as e:
                self.tw = None
                messagebox.showerror("FTP Lite", f"파일 전송 창을 열지 못했습니다.\n{e}", parent=self)
                return
        self.tw.toggle()

    def show_transfer(self):
        if self.tw is None or not self.tw.winfo_viewable():
            self.toggle_transfer()

    # ---- 트레이 / 단축키에서 온 명령 처리 ----
    def _poll_events(self):
        try:
            while True:
                ev = self.events.get_nowait()
                if ev in ("toggle", "transfer"):
                    self.toggle_transfer()
                elif ev == "show":
                    self.show_transfer()
                elif ev == "settings":
                    self.open_settings()
                elif ev == "quit":
                    self.quit_app()
                    return
                elif isinstance(ev, tuple) and ev[0] == "error":
                    messagebox.showwarning("FTP Lite", ev[1])
        except queue.Empty:
            pass
        self.after(100, self._poll_events)

    # ---- 설정 창 ----
    def open_settings(self):
        if self.settings_win is not None and self.settings_win.winfo_exists():
            self.settings_win.deiconify()
            self.settings_win.lift()
            self.settings_win.focus_force()
            return
        self.settings_win = SiteManager(self, on_apply=self._settings_applied)
        self.settings_win.lift()
        self.settings_win.focus_force()

    def _settings_applied(self):
        first = self._first_run_pending
        self._first_run_pending = False
        self.app_cfg = load_app_cfg()          # 설정 창에서 바꾼 값을 다시 읽는다
        if tw_hotkey_cfg(self.app_cfg) != self.tw_hotkey:
            self.set_tw_hotkey(tw_hotkey_cfg(self.app_cfg))
        if self.tw is not None:
            self.tw.on_settings()
        if first:                                # 첫 설정을 마치면 그때 파일 전송 창을 보여 준다
            self.after(300, self.show_transfer)

    def quit_app(self):
        if self._closing:
            return
        self._closing = True
        self._save_cfg()
        if self.tw is not None:
            self.tw.shutdown()
        if self.tray:
            self.tray.stop()
        self.destroy()


if __name__ == "__main__":
    if IS_WINDOWS:
        try:        # 작업표시줄이 pythonw.exe(파이썬 아이콘)와 묶지 않고 이 프로그램 고유 아이콘을 쓰게 한다
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("SungCheol.FTPLite.1")
        except Exception:
            pass
    if single_instance():
        LiteApp().mainloop()
    elif not show_running_instance():
        _root = tk.Tk()
        _root.withdraw()
        messagebox.showinfo("FTP Lite", "이미 실행 중입니다.\n트레이 아이콘을 누르거나 설정한 단축키(기본 Ctrl+Alt+F)로 창을 불러오세요.")
        _root.destroy()
