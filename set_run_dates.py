import os
import sys
import requests
import urllib3
import tkinter as tk
from tkinter import ttk, messagebox
import calendar
from datetime import datetime, date, timedelta, timezone
from dotenv import load_dotenv

__author__ = "junghun.lee"

load_dotenv()
urllib3.disable_warnings()

# ══════════════════════════════════════════════
# 공통 설정
# ══════════════════════════════════════════════
BASE_URL   = os.getenv("TESTRAIL_URL", "").rstrip("/")
AUTH       = (os.getenv("TESTRAIL_USER"), os.getenv("TESTRAIL_API_KEY"))
PROJECT_ID = int(os.getenv("TESTRAIL_PROJECT_ID", "0"))

if BASE_URL and not BASE_URL.startswith("http"):
    BASE_URL = "https://" + BASE_URL


def api_get(endpoint):
    res = requests.get(
        f"{BASE_URL}/index.php?/api/v2/{endpoint}",
        auth=AUTH, verify=False
    )
    if res.status_code != 200:
        return {}
    return res.json()


# [fix6] 페이지네이션 헬퍼 -- limit/offset 으로 전체 레코드를 순차 로드
def api_get_paged(base_endpoint, list_key, limit=250):
    results = []
    offset  = 0
    while True:
        data  = api_get(f"{base_endpoint}&limit={limit}&offset={offset}")
        items = data.get(list_key, [])
        results.extend(items)
        if len(items) < limit:
            break
        offset += limit
    return results


def api_post(endpoint, payload=None):
    res = requests.post(
        f"{BASE_URL}/index.php?/api/v2/{endpoint}",
        auth=AUTH, verify=False, json=payload or {}
    )
    return res


# [fix5] UTC 기준 날짜 변환 헬퍼
def ts_to_date_str(ts):
    """Unix timestamp → YYYY-MM-DD (UTC)"""
    return datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d")


def date_str_to_ts(date_str):
    """YYYY-MM-DD → Unix timestamp (UTC midnight)"""
    return int(datetime.strptime(date_str, "%Y-%m-%d")
               .replace(tzinfo=timezone.utc).timestamp())


# ══════════════════════════════════════════════
# 달력 팝업 위젯  (라이트 테마)
# ══════════════════════════════════════════════
class CalendarPopup(tk.Toplevel):

    BG         = "#ffffff"
    HEADER_BG  = "#4A90D9"
    HEADER_FG  = "#ffffff"
    TODAY_BG   = "#FFF3CD"
    TODAY_FG   = "#856404"
    SEL_BG     = "#4A90D9"
    SEL_FG     = "#ffffff"
    DAY_FG     = "#212529"
    WEEKEND_FG = "#dc3545"
    DOW_FG     = "#6c757d"
    HOVER_BG   = "#e7f0fb"
    BTN_BG     = "#f8f9fa"
    BTN_FG     = "#4A90D9"
    BORDER     = "#dee2e6"

    def __init__(self, parent, title="Select Date", initial_date=None):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.grab_set()
        self.configure(bg=self.BG)

        self.selected_date = None
        today  = date.today()
        init   = initial_date or today
        self._year   = init.year
        self._month  = init.month
        self._today  = today

        self._build_ui()
        self._render_calendar()
        self._center(parent)

    def _build_ui(self):
        hdr = tk.Frame(self, bg=self.HEADER_BG, padx=10, pady=8)
        hdr.pack(fill="x")

        arrow_kw = dict(
            bg=self.HEADER_BG, fg=self.HEADER_FG,
            font=("Segoe UI", 14, "bold"),
            relief="flat", bd=0, cursor="hand2",
            activebackground="#357ABD",
            activeforeground=self.HEADER_FG,
            padx=8, pady=0
        )
        tk.Button(hdr, text="<", command=self._prev_month, **arrow_kw).pack(side="left")
        tk.Button(hdr, text=">", command=self._next_month, **arrow_kw).pack(side="right")

        self._lbl_month = tk.Label(
            hdr, bg=self.HEADER_BG, fg=self.HEADER_FG,
            font=("Segoe UI", 12, "bold")
        )
        self._lbl_month.pack(side="left", expand=True)

        dow_fr = tk.Frame(self, bg="#f8f9fa", pady=5)
        dow_fr.pack(fill="x")
        for i, d in enumerate(["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"]):
            fg = self.WEEKEND_FG if i in (0, 6) else self.DOW_FG
            tk.Label(
                dow_fr, text=d, width=4, bg="#f8f9fa", fg=fg,
                font=("Segoe UI", 9, "bold")
            ).grid(row=0, column=i, padx=2)

        tk.Frame(self, bg=self.BORDER, height=1).pack(fill="x")

        self._cal_frame = tk.Frame(self, bg=self.BG, padx=8, pady=6)
        self._cal_frame.pack()

        self._day_btns = []
        for r in range(6):
            row = []
            for c in range(7):
                lbl = tk.Label(
                    self._cal_frame, width=4, height=1,
                    bg=self.BG, fg=self.DAY_FG,
                    font=("Segoe UI", 11),
                    relief="flat", cursor="hand2", bd=0
                )
                lbl.grid(row=r, column=c, padx=2, pady=2)
                lbl.bind("<Button-1>", self._on_day_click)
                lbl.bind("<Enter>",    self._on_hover_enter)
                lbl.bind("<Leave>",    self._on_hover_leave)
                row.append(lbl)
            self._day_btns.append(row)

        tk.Frame(self, bg=self.BORDER, height=1).pack(fill="x")

        foot = tk.Frame(self, bg=self.BTN_BG, pady=6)
        foot.pack(fill="x")
        tk.Button(
            foot, text="Today", command=self._go_today,
            bg=self.BTN_BG, fg=self.BTN_FG,
            font=("Segoe UI", 9, "bold"),
            relief="flat", bd=0, cursor="hand2",
            activebackground=self.HOVER_BG,
            padx=12, pady=3
        ).pack()

    def _render_calendar(self):
        self._lbl_month.config(text=f"{self._year}.{self._month:02d}")

        calendar.setfirstweekday(6)
        cal_matrix = calendar.monthcalendar(self._year, self._month)
        while len(cal_matrix) < 6:
            cal_matrix.append([0] * 7)

        for r, week in enumerate(cal_matrix):
            for c, day in enumerate(week):
                lbl = self._day_btns[r][c]
                if day == 0:
                    lbl.config(text="", bg=self.BG, state="disabled",
                               cursor="arrow", fg="#cccccc")
                    lbl._day = None
                else:
                    is_today = (day == self._today.day and
                                self._month == self._today.month and
                                self._year  == self._today.year)
                    is_sel   = (self.selected_date and
                                day == self.selected_date.day and
                                self._month == self.selected_date.month and
                                self._year  == self.selected_date.year)
                    is_we    = c in (0, 6)

                    if is_sel:
                        bg, fg, font = self.SEL_BG, self.SEL_FG, ("Segoe UI", 11, "bold")
                    elif is_today:
                        bg, fg, font = self.TODAY_BG, self.TODAY_FG, ("Segoe UI", 11, "bold")
                    else:
                        bg   = self.BG
                        fg   = self.WEEKEND_FG if is_we else self.DAY_FG
                        font = ("Segoe UI", 11)

                    lbl.config(text=str(day), bg=bg, fg=fg,
                               font=font, state="normal", cursor="hand2")
                    lbl._day = day

    def _on_day_click(self, event):
        lbl = event.widget
        if not getattr(lbl, "_day", None):
            return
        self.selected_date = date(self._year, self._month, lbl._day)
        self.destroy()

    def _on_hover_enter(self, event):
        lbl = event.widget
        if not getattr(lbl, "_day", None):
            return
        is_sel = (self.selected_date and
                  lbl._day == self.selected_date.day and
                  self._month == self.selected_date.month and
                  self._year  == self.selected_date.year)
        if not is_sel:
            lbl.config(bg=self.HOVER_BG)

    def _on_hover_leave(self, event):
        lbl = event.widget
        if not getattr(lbl, "_day", None):
            return
        is_sel   = (self.selected_date and
                    lbl._day == self.selected_date.day and
                    self._month == self.selected_date.month and
                    self._year  == self.selected_date.year)
        is_today = (lbl._day == self._today.day and
                    self._month == self._today.month and
                    self._year  == self._today.year)
        if is_sel:
            lbl.config(bg=self.SEL_BG)
        elif is_today:
            lbl.config(bg=self.TODAY_BG)
        else:
            lbl.config(bg=self.BG)

    def _prev_month(self):
        if self._month == 1:
            self._month, self._year = 12, self._year - 1
        else:
            self._month -= 1
        self._render_calendar()

    def _next_month(self):
        if self._month == 12:
            self._month, self._year = 1, self._year + 1
        else:
            self._month += 1
        self._render_calendar()

    def _go_today(self):
        self._year  = self._today.year
        self._month = self._today.month
        self._render_calendar()

    def _center(self, parent):
        self.update_idletasks()
        pw = parent.winfo_rootx() + parent.winfo_width()  // 2
        ph = parent.winfo_rooty() + parent.winfo_height() // 2
        w  = self.winfo_width()
        h  = self.winfo_height()
        self.geometry(f"+{pw - w//2}+{ph - h//2}")


# ══════════════════════════════════════════════
# 메인 앱  (라이트 테마)
# ══════════════════════════════════════════════
class App(tk.Tk):

    BG        = "#f4f6f9"
    SURFACE   = "#ffffff"
    HEADER_BG = "#4A90D9"
    HEADER_FG = "#ffffff"
    TEXT      = "#212529"
    SUBTEXT   = "#6c757d"
    ACCENT    = "#4A90D9"
    ACCENT_FG = "#ffffff"
    GREEN     = "#198754"
    RED       = "#dc3545"
    BORDER    = "#dee2e6"
    ROW_ALT   = "#EEF4FF"
    HDR_BG    = "#D6E4F7"
    BTN_GHOST = "#e9ecef"
    TOOLBAR   = "#f0f4f8"
    ROW_COLORS = [
        "#4A90D9", "#34A853", "#E67E22", "#9B59B6",
        "#E91E63", "#00ACC1", "#F4511E", "#43A047",
    ]

    #        c0  c1  c2   c3   c4  c5   c6  c7   c8
    #       bar chk  no  name  sd  📅   ed  📅  status
    COL_W = [5,  28,  30, 220, 110, 30, 110, 30,  90]

    def __init__(self):
        super().__init__()
        self.title("TestRail — Plan Entry Date Setter  v3")
        self.configure(bg=self.BG)
        self.resizable(False, False)

        self._plans    = []
        self._entries  = []
        self._run_rows = []

        self._search_var     = tk.StringVar()
        self._auto_start_var = tk.StringVar()
        self._auto_days_var  = tk.StringVar(value="7")
        self._auto_gap_var   = tk.StringVar(value="0")

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TCombobox",
            fieldbackground=self.SURFACE,
            background=self.SURFACE,
            foreground=self.TEXT,
            selectbackground=self.ACCENT,
            selectforeground=self.ACCENT_FG,
            font=("Segoe UI", 10)
        )
        style.configure("TScrollbar",
            background=self.BTN_GHOST,
            troughcolor=self.BG,
            bordercolor=self.BORDER,
            arrowcolor=self.SUBTEXT
        )

        self.geometry("860x680")
        self._build_ui()
        self._load_plans()
        self._center()

    def _build_ui(self):
        hdr = tk.Frame(self, bg=self.HEADER_BG, padx=20, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="TestRail  |  Plan Entry Date Setter",
                 bg=self.HEADER_BG, fg=self.HEADER_FG,
                 font=("Segoe UI", 13, "bold")).pack(side="left")
        tk.Label(hdr, text=f"Project {PROJECT_ID}   {BASE_URL}",
                 bg=self.HEADER_BG, fg="#cce0f5",
                 font=("Segoe UI", 9)).pack(side="right")

        plan_card = tk.Frame(self, bg=self.SURFACE,
                             padx=20, pady=12,
                             highlightbackground=self.BORDER,
                             highlightthickness=1)
        plan_card.pack(fill="x", padx=16, pady=(14, 6))

        tk.Label(plan_card, text="Test Plan",
                 bg=self.SURFACE, fg=self.SUBTEXT,
                 font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 12))

        self._plan_var = tk.StringVar()
        self._plan_cb  = ttk.Combobox(plan_card, textvariable=self._plan_var,
                                       state="readonly", width=50,
                                       font=("Segoe UI", 10))
        self._plan_cb.grid(row=0, column=1, sticky="w")
        self._plan_cb.bind("<<ComboboxSelected>>", self._on_plan_selected)

        self._lbl_loading = tk.Label(plan_card, text="Loading...",
                                      bg=self.SURFACE, fg=self.SUBTEXT,
                                      font=("Segoe UI", 9))
        self._lbl_loading.grid(row=0, column=2, padx=(12, 0))

        # ── 툴바 (Select All / Deselect All / Bulk Date) ────
        toolbar = tk.Frame(self, bg=self.TOOLBAR,
                           padx=16, pady=7,
                           highlightbackground=self.BORDER,
                           highlightthickness=1)
        toolbar.pack(fill="x", padx=16, pady=(0, 0))

        btn_kw = dict(relief="flat", bd=0, cursor="hand2",
                      font=("Segoe UI", 9),
                      padx=10, pady=4)

        tk.Button(toolbar, text="Select All",
                  command=self._select_all,
                  bg="#4A90D9", fg="white",
                  activebackground="#357ABD", activeforeground="white",
                  **btn_kw).pack(side="left", padx=(0, 4))

        tk.Button(toolbar, text="Deselect All",
                  command=self._deselect_all,
                  bg=self.BTN_GHOST, fg=self.TEXT,
                  activebackground=self.BORDER, activeforeground=self.TEXT,
                  **btn_kw).pack(side="left", padx=(0, 16))

        tk.Frame(toolbar, bg=self.BORDER, width=1).pack(side="left", fill="y", padx=(0, 12))

        tk.Label(toolbar, text="Bulk Date:",
                 bg=self.TOOLBAR, fg=self.SUBTEXT,
                 font=("Segoe UI", 9, "bold")).pack(side="left")

        self._bulk_start_var = tk.StringVar()
        tk.Entry(toolbar, textvariable=self._bulk_start_var,
                 width=11, bg=self.SURFACE, fg=self.TEXT,
                 relief="solid", bd=1,
                 font=("Segoe UI", 9),
                 state="readonly"
                 ).pack(side="left", padx=(6, 2))
        tk.Button(toolbar, text="📅",
                  command=lambda: self._pick_date(self._bulk_start_var, "Bulk Start Date"),
                  bg=self.TOOLBAR, fg="#4A90D9",
                  relief="flat", bd=0, cursor="hand2",
                  font=("Segoe UI", 10),
                  activebackground=self.BTN_GHOST
                  ).pack(side="left", padx=(0, 8))

        tk.Label(toolbar, text="~",
                 bg=self.TOOLBAR, fg=self.SUBTEXT,
                 font=("Segoe UI", 10)).pack(side="left")

        self._bulk_end_var = tk.StringVar()
        tk.Entry(toolbar, textvariable=self._bulk_end_var,
                 width=11, bg=self.SURFACE, fg=self.TEXT,
                 relief="solid", bd=1,
                 font=("Segoe UI", 9),
                 state="readonly"
                 ).pack(side="left", padx=(8, 2))
        tk.Button(toolbar, text="📅",
                  command=lambda: self._pick_date(self._bulk_end_var, "Bulk End Date"),
                  bg=self.TOOLBAR, fg="#4A90D9",
                  relief="flat", bd=0, cursor="hand2",
                  font=("Segoe UI", 10),
                  activebackground=self.BTN_GHOST
                  ).pack(side="left", padx=(0, 8))

        tk.Button(toolbar, text="Apply to Checked",
                  command=self._apply_bulk_dates,
                  bg="#34A853", fg="white",
                  activebackground="#2d8f47", activeforeground="white",
                  **btn_kw).pack(side="left")

        # ── 툴바2: Auto-Date + 검색 ──
        toolbar2 = tk.Frame(self, bg=self.TOOLBAR,
                            padx=16, pady=7,
                            highlightbackground=self.BORDER,
                            highlightthickness=1)
        toolbar2.pack(fill="x", padx=16, pady=(0, 2))

        tk.Label(toolbar2, text="Auto-Date:",
                 bg=self.TOOLBAR, fg=self.SUBTEXT,
                 font=("Segoe UI", 9, "bold")).pack(side="left")
        tk.Entry(toolbar2, textvariable=self._auto_start_var,
                 width=11, bg=self.SURFACE, fg=self.TEXT,
                 relief="solid", bd=1, font=("Segoe UI", 9), state="readonly"
                 ).pack(side="left", padx=(6, 2))
        tk.Button(toolbar2, text="📅",
                  command=lambda: self._pick_date(self._auto_start_var, "Auto-Date: Start"),
                  bg=self.TOOLBAR, fg="#4A90D9", relief="flat", bd=0, cursor="hand2",
                  font=("Segoe UI", 10), activebackground=self.BTN_GHOST
                  ).pack(side="left", padx=(0, 10))
        tk.Label(toolbar2, text="Days:", bg=self.TOOLBAR, fg=self.SUBTEXT,
                 font=("Segoe UI", 9)).pack(side="left")
        tk.Entry(toolbar2, textvariable=self._auto_days_var,
                 width=4, bg=self.SURFACE, fg=self.TEXT,
                 relief="solid", bd=1, font=("Segoe UI", 9)
                 ).pack(side="left", padx=(4, 10))
        tk.Label(toolbar2, text="Gap:", bg=self.TOOLBAR, fg=self.SUBTEXT,
                 font=("Segoe UI", 9)).pack(side="left")
        tk.Entry(toolbar2, textvariable=self._auto_gap_var,
                 width=4, bg=self.SURFACE, fg=self.TEXT,
                 relief="solid", bd=1, font=("Segoe UI", 9)
                 ).pack(side="left", padx=(4, 10))
        tk.Button(toolbar2, text="Auto Assign",
                  command=self._auto_assign_dates,
                  bg="#E67E22", fg="white",
                  activebackground="#CA6F1E", activeforeground="white",
                  **btn_kw).pack(side="left", padx=(0, 16))
        tk.Frame(toolbar2, bg=self.BORDER, width=1).pack(side="left", fill="y", padx=(0, 12))
        tk.Label(toolbar2, text="🔍", bg=self.TOOLBAR, fg=self.SUBTEXT,
                 font=("Segoe UI", 10)).pack(side="left")
        tk.Entry(toolbar2, textvariable=self._search_var,
                 width=22, bg=self.SURFACE, fg=self.TEXT,
                 relief="solid", bd=1, font=("Segoe UI", 9)
                 ).pack(side="left", padx=(4, 6))
        self._lbl_filter_count = tk.Label(toolbar2, text="",
                 bg=self.TOOLBAR, fg=self.SUBTEXT, font=("Segoe UI", 9))
        self._lbl_filter_count.pack(side="left")
        self._search_var.trace_add("write", self._filter_runs)

        # Run 테이블
        tbl_wrap = tk.Frame(self, bg=self.BG, padx=16, pady=4)
        tbl_wrap.pack(fill="both", expand=True)

        hdr_row = tk.Frame(tbl_wrap, bg=self.HDR_BG,
                           highlightbackground=self.BORDER,
                           highlightthickness=1)
        hdr_row.pack(fill="x")

        for c, (txt, px) in enumerate(zip(
            ["", "", "", "Suite / Run Name",
             "Start Date", "", "End Date", "", "Status"],
            self.COL_W
        )):
            hdr_row.columnconfigure(c, minsize=px)
            tk.Label(hdr_row, text=txt,
                     bg=self.HDR_BG, fg=self.SUBTEXT,
                     font=("Segoe UI", 9, "bold"),
                     anchor="w", padx=6, pady=7
                     ).grid(row=0, column=c, sticky="ew")

        canvas_wrap = tk.Frame(tbl_wrap, bg=self.SURFACE,
                               highlightbackground=self.BORDER,
                               highlightthickness=1)
        canvas_wrap.pack(fill="both", expand=True)

        self._canvas = tk.Canvas(canvas_wrap, bg=self.SURFACE,
                                  highlightthickness=0, height=340)
        self._scrollbar = ttk.Scrollbar(canvas_wrap, orient="vertical",
                                        command=self._canvas.yview)
        self._scroll_fr = tk.Frame(self._canvas, bg=self.SURFACE)

        self._scroll_fr.bind("<Configure>", self._on_scroll_frame_configure)
        self._canvas.create_window((0, 0), window=self._scroll_fr, anchor="nw")
        self._canvas.bind(
            "<Configure>",
            lambda e: self._canvas.itemconfig(
                self._canvas.find_all()[0], width=e.width
            )
        )
        self._canvas.configure(yscrollcommand=self._scrollbar.set)
        self._canvas.pack(side="left", fill="both", expand=True)
        self._scrollbar.pack(side="right", fill="y")
        self._canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        foot = tk.Frame(self, bg=self.SURFACE,
                        padx=16, pady=10,
                        highlightbackground=self.BORDER,
                        highlightthickness=1)
        foot.pack(fill="x", padx=16, pady=(4, 14))

        self._lbl_status = tk.Label(foot, text="",
                                     bg=self.SURFACE, fg=self.SUBTEXT,
                                     font=("Segoe UI", 9))
        self._lbl_status.pack(side="left")

        tk.Button(foot, text="Save Selected Runs",
                  command=self._save_dates,
                  bg=self.ACCENT, fg=self.ACCENT_FG,
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=20, pady=7,
                  cursor="hand2",
                  activebackground="#357ABD",
                  activeforeground=self.ACCENT_FG, bd=0
                  ).pack(side="right", padx=(6, 0))

        tk.Button(foot, text="Refresh",
                  command=self._load_plans,
                  bg=self.BTN_GHOST, fg=self.TEXT,
                  font=("Segoe UI", 10),
                  relief="flat", padx=14, pady=7,
                  cursor="hand2",
                  activebackground=self.BORDER,
                  activeforeground=self.TEXT, bd=0
                  ).pack(side="right")

    # ── 데이터 로딩 ───────────────────────────
    def _load_plans(self):
        self._lbl_loading.config(text="Loading...", fg=self.SUBTEXT)
        self.update()
        try:
            # [fix6] 페이지네이션 적용
            raw_plans = api_get_paged(
                f"get_plans/{PROJECT_ID}&is_completed=0", "plans")
            ms_list   = api_get_paged(
                f"get_milestones/{PROJECT_ID}&is_completed=0", "milestones")
        except Exception as e:
            messagebox.showerror("Error", f"API connection failed:\n{e}")
            self._lbl_loading.config(text="Connection failed", fg=self.RED)
            return

        ms_map = {}
        for m in ms_list:
            ms_map[m["id"]] = m["name"]
            for sub in m.get("milestones", []):
                ms_map[sub["id"]] = f"{m['name']} > {sub['name']}"

        combined = []
        for p in raw_plans:
            ms_name = ms_map.get(p.get("milestone_id"), "")
            ms_tag  = f"  [{ms_name}]" if ms_name else ""
            combined.append({
                "label" : f"[Plan] {p['name']}{ms_tag}",
                "type"  : "plan",
                "id"    : p["id"],
                "name"  : p["name"],
            })

        ms_run_count      = 0
        all_milestone_ids = []
        for m in ms_list:
            all_milestone_ids.append(m["id"])
            for sub in m.get("milestones", []):
                all_milestone_ids.append(sub["id"])

        seen_run_ids = set()
        for ms_id in all_milestone_ids:
            try:
                # [fix6] 페이지네이션 적용
                runs = api_get_paged(
                    f"get_runs/{PROJECT_ID}&milestone_id={ms_id}", "runs")
                for r in runs:
                    if r["id"] in seen_run_ids:
                        continue
                    seen_run_ids.add(r["id"])
                    ms_label = ms_map.get(ms_id, "")
                    ms_tag   = f"  [{ms_label}]" if ms_label else ""
                    combined.append({
                        "label"        : f"[Run]  {r['name']}{ms_tag}",
                        "type"         : "run",
                        "id"           : r["id"],
                        "name"         : r["name"],
                        "milestone_id" : ms_id,
                    })
                    ms_run_count += 1
            except Exception:
                pass

        # 독립 Test Run (Plan/Milestone 미포함)
        try:
            # [fix6] 페이지네이션 적용
            all_runs = api_get_paged(f"get_runs/{PROJECT_ID}", "runs")
            for r in all_runs:
                if r["id"] in seen_run_ids:
                    continue
                seen_run_ids.add(r["id"])
                ms_id    = r.get("milestone_id")
                ms_label = ms_map.get(ms_id, "") if ms_id else ""
                ms_tag   = f"  [{ms_label}]" if ms_label else ""
                combined.append({
                    "label"        : f"[Run]  {r['name']}{ms_tag}  (ID: {r['id']})",
                    "type"         : "run",
                    "id"           : r["id"],
                    "name"         : r["name"],
                    "milestone_id" : ms_id,
                })
                ms_run_count += 1
        except Exception:
            pass

        self._plans = combined
        labels = [c["label"] for c in combined]

        # [fix7] Refresh 후 현재 선택 플랜 유지
        prev_label = self._plan_var.get()

        self._plan_cb["values"] = labels
        # [fix8] 긴 플랜 이름이 짤리지 않도록 가장 긴 라벨 길이에 맞춰 너비 자동 조정
        if labels:
            max_w = max(len(l) for l in labels)
            self._plan_cb.configure(width=min(max_w, 100))
        total = len(combined)

        if total == 0:
            self._lbl_loading.config(
                text="No plans or runs found. Check PROJECT_ID in .env",
                fg=self.RED
            )
            return

        self._lbl_loading.config(
            text=f"{total} found  (plans: {len(raw_plans)}, runs: {ms_run_count})",
            fg=self.GREEN
        )

        if prev_label and prev_label in labels:
            self._plan_cb.current(labels.index(prev_label))
        else:
            self._plan_cb.current(0)
        self._on_plan_selected()

    def _on_plan_selected(self, event=None):
        idx = self._plan_cb.current()
        if idx < 0:
            return
        selected = self._plans[idx]
        if selected["type"] == "plan":
            self._load_entries_from_plan(selected["id"])
        else:
            self._load_entries_from_run(selected)

    def _load_entries_from_plan(self, plan_id):
        self._lbl_status.config(text="Loading runs...", fg=self.SUBTEXT)
        self.update()
        try:
            plan_detail = api_get(f"get_plan/{plan_id}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load plan:\n{e}")
            return

        self._entries = []
        for entry in plan_detail.get("entries", []):
            for run in entry.get("runs", []):
                self._entries.append({
                    "entry_id" : entry["id"],
                    "run_id"   : run["id"],
                    "name"     : entry.get("name") or run.get("name", ""),
                    "plan_id"  : plan_id,
                    "start_on" : run.get("start_on"),
                    "due_on"   : run.get("due_on"),
                    "source"   : "plan",
                })

        self._render_runs()
        self._lbl_status.config(
            text=f"{len(self._entries)} runs loaded", fg=self.GREEN
        )

    def _load_entries_from_run(self, run_item):
        self._lbl_status.config(text="Loading run...", fg=self.SUBTEXT)
        self.update()
        try:
            run_detail = api_get(f"get_run/{run_item['id']}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load run:\n{e}")
            return

        self._entries = [{
            "entry_id" : None,
            "run_id"   : run_detail.get("id"),
            "name"     : run_detail.get("name", run_item["name"]),
            "plan_id"  : None,
            "start_on" : run_detail.get("started_on") or run_detail.get("start_on"),
            "due_on"   : run_detail.get("due_on"),
            "source"   : "run",
        }]

        self._render_runs()
        self._lbl_status.config(text="1 run loaded  (Milestone Run)", fg=self.GREEN)

    def _load_entries(self, plan_id):
        self._load_entries_from_plan(plan_id)

    # ── Run 행 렌더링 ─────────────────────────
    def _render_runs(self):
        # [fix1] 이전 StringVar 트레이스 명시적 제거
        for row in self._run_rows:
            for var, mode, tid in row.get("_trace_ids", []):
                try:
                    var.trace_remove(mode, tid)
                except Exception:
                    pass

        for w in self._scroll_fr.winfo_children():
            w.destroy()
        self._run_rows.clear()

        if not self._entries:
            tk.Label(self._scroll_fr, text="No runs found.",
                     bg=self.SURFACE, fg=self.SUBTEXT,
                     font=("Segoe UI", 10), pady=24).pack()
            self._lbl_filter_count.config(text="0", fg=self.SUBTEXT)
            return

        for i, entry in enumerate(self._entries):
            row_bg    = self.SURFACE if i % 2 == 0 else self.ROW_ALT
            bar_color = self.ROW_COLORS[i % len(self.ROW_COLORS)]

            divider_fr = None
            if i > 0:
                divider_fr = tk.Frame(self._scroll_fr, bg=self.BORDER, height=1)
                divider_fr.pack(fill="x")

            row_fr = tk.Frame(self._scroll_fr, bg=row_bg)
            row_fr.pack(fill="x")
            for _c, _px in enumerate(self.COL_W):
                row_fr.columnconfigure(_c, minsize=_px)

            chk_var = tk.BooleanVar(value=True)

            # c0: 좌측 컬러 인디케이터 바
            tk.Frame(row_fr, bg=bar_color, width=5).grid(
                row=0, column=0, sticky="ns", ipady=14)

            # c1: 체크박스
            tk.Checkbutton(row_fr, variable=chk_var,
                           bg=row_bg, activebackground=row_bg,
                           cursor="hand2"
                           ).grid(row=0, column=1, padx=(4, 0), pady=8)

            # c2: 번호 배지
            tk.Label(row_fr, text=f"{i+1:02d}",
                     bg=bar_color, fg="white",
                     font=("Segoe UI", 8, "bold"),
                     width=3, padx=2, pady=2
                     ).grid(row=0, column=2, padx=(4, 0))

            # c3: Run 이름
            tk.Label(row_fr, text=entry["name"],
                     bg=row_bg, fg=self.TEXT,
                     font=("Segoe UI", 10, "bold"), anchor="w",
                     padx=8).grid(row=0, column=3, sticky="ew")

            # c4: Start Date 입력창
            start_var = tk.StringVar()
            if entry.get("start_on"):
                try:
                    start_var.set(ts_to_date_str(entry["start_on"]))  # [fix5]
                except Exception:
                    pass

            tk.Entry(row_fr, textvariable=start_var,
                     width=12, bg=self.SURFACE, fg=self.TEXT,
                     disabledbackground="#EEF4FF",
                     disabledforeground=self.TEXT,
                     relief="solid", bd=1,
                     font=("Segoe UI", 10),
                     state="readonly"
                     ).grid(row=0, column=4, padx=(4, 2))

            # c5: Start Date 달력 버튼
            tk.Button(row_fr, text="📅",
                      command=lambda sv=start_var: self._pick_date(sv, "Start Date"),
                      bg=row_bg, fg=self.ACCENT,
                      relief="flat", bd=0, cursor="hand2",
                      font=("Segoe UI", 11),
                      activebackground=self.BTN_GHOST
                      ).grid(row=0, column=5, padx=(0, 10))

            # c6: End Date 입력창
            end_var = tk.StringVar()
            if entry.get("due_on"):
                try:
                    end_var.set(ts_to_date_str(entry["due_on"]))  # [fix5]
                except Exception:
                    pass

            tk.Entry(row_fr, textvariable=end_var,
                     width=12, bg=self.SURFACE, fg=self.TEXT,
                     disabledbackground="#EEF4FF",
                     disabledforeground=self.TEXT,
                     relief="solid", bd=1,
                     font=("Segoe UI", 10),
                     state="readonly"
                     ).grid(row=0, column=6, padx=(4, 2))

            # c7: End Date 달력 버튼
            tk.Button(row_fr, text="📅",
                      command=lambda ev=end_var: self._pick_date(ev, "End Date"),
                      bg=row_bg, fg=self.ACCENT,
                      relief="flat", bd=0, cursor="hand2",
                      font=("Segoe UI", 11),
                      activebackground=self.BTN_GHOST
                      ).grid(row=0, column=7, padx=(0, 8))

            # c8: 저장 상태 라벨
            status_lbl = tk.Label(row_fr, text="",
                                   bg=row_bg, fg=self.SUBTEXT,
                                   font=("Segoe UI", 9), width=10, anchor="w")
            status_lbl.grid(row=0, column=8, padx=(2, 8))

            row_data = {
                "entry":      entry,
                "chk_var":    chk_var,
                "start_var":  start_var,
                "end_var":    end_var,
                "status_lbl": status_lbl,
                "row_fr":     row_fr,
                "divider":    divider_fr,
                "_trace_ids": [],
            }
            self._run_rows.append(row_data)

            # [fix1] row_data 완성 후 트레이스 등록
            tid_s = start_var.trace_add("write", lambda *a, r=row_data: self._validate_row(r))
            tid_e = end_var.trace_add("write",   lambda *a, r=row_data: self._validate_row(r))
            row_data["_trace_ids"] = [
                (start_var, "write", tid_s),
                (end_var,   "write", tid_e),
            ]
            self._validate_row(row_data)

        self._search_var.set("")

    # ── Select All / Deselect All ────────────
    def _select_all(self):
        for row in self._run_rows:
            if row["row_fr"].winfo_ismapped():
                row["chk_var"].set(True)

    def _deselect_all(self):
        for row in self._run_rows:
            if row["row_fr"].winfo_ismapped():
                row["chk_var"].set(False)

    # ── Bulk Date 일괄 적용 ───────────────────
    def _apply_bulk_dates(self):
        s = self._bulk_start_var.get()
        e = self._bulk_end_var.get()

        if not s and not e:
            messagebox.showwarning("Warning",
                "Please select at least one date in the Bulk Date fields.")
            return

        if s and e:
            try:
                if datetime.strptime(s, "%Y-%m-%d") > datetime.strptime(e, "%Y-%m-%d"):
                    messagebox.showerror("Date Error",
                        "Bulk Start Date is after End Date.")
                    return
            except ValueError:
                pass

        targets = [r for r in self._run_rows
                   if r["chk_var"].get() and r["row_fr"].winfo_ismapped()]
        if not targets:
            messagebox.showwarning("Warning", "No runs selected (checked).")
            return

        applied = 0
        for row in targets:
            if s:
                row["start_var"].set(s)
            if e:
                row["end_var"].set(e)
            applied += 1

        self._lbl_status.config(
            text=f"Bulk applied to {applied} run(s)  —  click Save to confirm",
            fg="#856404"
        )

    # =============================================
    # v2 신규 메서드
    # =============================================

    def _validate_row(self, row):
        s = row["start_var"].get()
        e = row["end_var"].get()
        lbl = row.get("status_lbl")
        if lbl is None:
            return
        if s and e:
            try:
                if datetime.strptime(s, "%Y-%m-%d") > datetime.strptime(e, "%Y-%m-%d"):
                    lbl.config(text="⚠ Start>End", fg=self.RED)
                    return
            except ValueError:
                pass
        if lbl.cget("text") == "⚠ Start>End":
            lbl.config(text="", fg=self.SUBTEXT)

    def _collect_invalid_rows(self, targets):
        errors = []
        for row in targets:
            s = row["start_var"].get()
            e = row["end_var"].get()
            if s and e:
                try:
                    if datetime.strptime(s, "%Y-%m-%d") > datetime.strptime(e, "%Y-%m-%d"):
                        errors.append(
                            f"'{row['entry']['name']}': Start Date is after End Date")
                except ValueError:
                    errors.append(
                        f"'{row['entry']['name']}': Invalid date format")
        return errors

    def _filter_runs(self, *args):
        if not self._run_rows:
            self._lbl_filter_count.config(text="", fg=self.SUBTEXT)
            return
        query     = self._search_var.get().strip().lower()
        filtering = (query != "")
        for row in self._run_rows:
            row["row_fr"].pack_forget()
            if row.get("divider") is not None:
                row["divider"].pack_forget()
        visible       = 0
        first_visible = True
        for row in self._run_rows:
            matched = not filtering or (query in row["entry"]["name"].lower())
            if matched:
                divider = row.get("divider")
                if divider is not None and not first_visible and not filtering:
                    divider.pack(fill="x")
                row["row_fr"].pack(fill="x")
                first_visible = False
                visible += 1
        total = len(self._run_rows)
        if filtering:
            self._lbl_filter_count.config(
                text=f"{visible}/{total}",
                fg=self.GREEN if visible > 0 else self.RED)
        else:
            self._lbl_filter_count.config(text=str(total), fg=self.SUBTEXT)

    def _auto_assign_dates(self):
        # [fix8] 메시지 영어로 통일
        start_str = self._auto_start_var.get()
        if not start_str:
            messagebox.showwarning("Warning",
                "Please select a Start Date for Auto-Date first.")
            return
        try:
            days = int(self._auto_days_var.get())
            if days < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Days must be an integer ≥ 1.")
            return
        try:
            gap = int(self._auto_gap_var.get())
            if gap < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Gap must be an integer ≥ 0.")
            return
        try:
            base_start = datetime.strptime(start_str, "%Y-%m-%d").date()
        except ValueError:
            messagebox.showerror("Error",
                "Invalid date format for Auto-Date Start (YYYY-MM-DD).")
            return

        targets = [r for r in self._run_rows
                   if r["chk_var"].get() and r["row_fr"].winfo_ismapped()]
        if not targets:
            messagebox.showwarning("Warning",
                "No checked runs visible in the current filter.")
            return

        cursor = base_start
        for row in targets:
            run_end = cursor + timedelta(days=days - 1)
            row["start_var"].set(cursor.strftime("%Y-%m-%d"))
            row["end_var"].set(run_end.strftime("%Y-%m-%d"))
            cursor = run_end + timedelta(days=gap + 1)

        self._lbl_status.config(
            text=f"Auto-Date: {len(targets)} run(s) assigned  —  click Save to confirm",
            fg="#856404")

    # ── 달력 팝업 ─────────────────────────────
    def _pick_date(self, str_var, title="Select Date"):
        init_date = None
        val = str_var.get()
        if val:
            try:
                init_date = datetime.strptime(val, "%Y-%m-%d").date()
            except ValueError:
                pass
        popup = CalendarPopup(self, title=title, initial_date=init_date)
        self.wait_window(popup)
        if popup.selected_date:
            str_var.set(popup.selected_date.strftime("%Y-%m-%d"))

    # ── 저장 ──────────────────────────────────
    def _save_dates(self):
        # [fix4] 가시 행만 대상
        targets = [r for r in self._run_rows
                   if r["chk_var"].get() and r["row_fr"].winfo_ismapped()]
        if not targets:
            messagebox.showwarning("Warning", "No runs selected.")
            return

        # [fix3] 날짜 역전 / 포맷 오류 사전 검사 후 차단
        errors = self._collect_invalid_rows(targets)
        if errors:
            messagebox.showerror(
                "Date Error",
                "Please fix the following before saving:\n\n" + "\n".join(errors)
            )
            return

        if not messagebox.askyesno("Confirm Save",
                                    f"Save dates for {len(targets)} run(s)?"):
            return

        success = fail = skip = 0
        fail_details = []

        for row in targets:
            entry     = row["entry"]
            start_str = row["start_var"].get()
            end_str   = row["end_var"].get()

            if not start_str and not end_str:
                row["status_lbl"].config(text="Skipped", fg=self.SUBTEXT)
                skip += 1
                continue

            payload = {}
            if start_str:
                payload["start_on"] = date_str_to_ts(start_str)  # [fix5]
            if end_str:
                payload["due_on"] = date_str_to_ts(end_str)       # [fix5]

            try:
                if entry.get("source") == "run" or not entry.get("entry_id"):
                    resp = api_post(f"update_run/{entry['run_id']}", payload)
                else:
                    resp = api_post(
                        f"update_plan_entry/{entry['plan_id']}/{entry['entry_id']}",
                        payload
                    )
                if resp.status_code == 200:
                    row["status_lbl"].config(text="Saved  ✓", fg=self.GREEN)
                    success += 1
                else:
                    try:
                        err_msg = resp.json().get("error", resp.text[:120])
                    except Exception:
                        err_msg = resp.text[:120]
                    row["status_lbl"].config(text="Failed  ✗", fg=self.RED)
                    fail_details.append(
                        f"[{entry['name']}]\n  HTTP {resp.status_code}: {err_msg}"
                    )
                    fail += 1
            except Exception as ex:
                row["status_lbl"].config(text="Error  ✗", fg=self.RED)
                fail_details.append(f"[{entry['name']}]\n  {ex}")
                fail += 1

        parts = []
        if success: parts.append(f"{success} saved")
        if skip:    parts.append(f"{skip} skipped")
        if fail:    parts.append(f"{fail} failed")
        self._lbl_status.config(
            text="  |  ".join(parts),
            fg=self.GREEN if fail == 0 else self.RED
        )

        if fail_details:
            messagebox.showerror("Save Failed", "\n\n".join(fail_details))

    # ── 스크롤 ────────────────────────────────
    def _on_scroll_frame_configure(self, event):
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_mousewheel(self, event):
        content_h = self._scroll_fr.winfo_reqheight()
        canvas_h  = self._canvas.winfo_height()
        if content_h <= canvas_h:
            return
        self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ── 화면 중앙 배치 ────────────────────────
    def _center(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = self.winfo_width()
        h  = self.winfo_height()
        self.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")


# ══════════════════════════════════════════════
# 실행
# ══════════════════════════════════════════════
if __name__ == "__main__":
    if not BASE_URL or not AUTH[0] or not AUTH[1] or PROJECT_ID == 0:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Configuration Error",
            "Please check your .env file:\n\n"
            "TESTRAIL_URL\nTESTRAIL_USER\nTESTRAIL_API_KEY\nTESTRAIL_PROJECT_ID"
        )
        sys.exit(1)

    app = App()
    app.mainloop()
