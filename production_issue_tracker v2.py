"""
AMAT Production Issue Tracker v2.0
Requires: pip install customtkinter openpyxl pillow
"""

import customtkinter as ctk
import tkinter as tk
import sys
from tkinter import messagebox, filedialog
import sqlite3, os, datetime, calendar
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image
from pilmoji import Pilmoji


# ══════════════════════════════════════════════════════════════════════════════
#  ✏️  EASY CUSTOMISATION — edit these blocks to change the look & feel
# ══════════════════════════════════════════════════════════════════════════════

# --- Colours -----------------------------------------------------------------
C = {
    "bg"      : "#e2ecf7",
    "surface" : "#f0f6fc",
    "panel"   : "#cddcee",
    "header"  : "#007ea5",
    "accent"  : "#2471c4",
    "accent_h": "#1a5ba0",
    "entry_bg": "#f5f9ff",
    "border"  : "#96b5d2",
    "text"    : "#0e1d2e",
    "subtle"  : "#557799",
    "open"    : "#b22222",
    "clarif"  : "#d87300",
    "closed"  : "#1a6e38",
    "stripe"  : "#e6f0fb",
    "sel"     : "#b5d2ee",
}

# --- Fonts -------------------------------------------------------------------
F = {
    "family" : "Montserrat",
    "size_sm": 10,
    "size_md": 11,
    "size_lg": 13,
    "size_xl": 15,
}

# --- Window ------------------------------------------------------------------
WIN_TITLE = "AMAT Production Issue Tracker v2.0"
WIN_SIZE  = "1420x820"
WIN_MIN   = (1000, 620)

# --- Table columns -----------------------------------------------------------
TABLE_COLS = [
    ("date",     "Date",             125, False),
    ("system",   "System No.",       230, True ),
    ("family",   "Product Family",   130, False),
    ("type",     "Issue Type",       165, True),
    ("sps",      "SPS No.",           145, False),
    ("ncr",      "NCR No.",          145, False),
    ("status",   "Status",            145, False),
    ("desc",     "Issue Description",530, True ),
    ("solution", "Solution",         530, True ),
    ("sol_date", "Sol. Date",         145, False),
    ("crf",      "CRF",               100, False),
    ("esw",      "ESW",               100, False),
    ("scv",      "SCV",               100, False),
    ("remarks",  "Remarks",          530, True ),
]

# --- Dropdown options --------------------------------------------------------
FAMILIES    = ["FEP", "DDP", "ETCH", "MDP", "EPI"]
ISSUE_TYPES = ["BOM Error", "Document Discrepancy", "Document Error",
               "Missing Document", "Design Error", "Others"]
TRACKER_TYPES = ["Engineering", "Material"]

# --- Paths -------------------------------------------------------------------
if getattr(sys, 'frozen', False):
    _HERE = os.path.dirname(sys.executable)
else:
    _HERE = os.path.dirname(os.path.abspath(__file__))

DB_PATH   = os.path.join(_HERE, "production_issues.db")
LOGO_PATH = os.path.join(_HERE, "logo.png")

# --- Derived filter lists (don't edit) ---------------------------------------
MONTH_NAMES = ["All"] + [datetime.date(2000, m, 1).strftime("%B") for m in range(1, 13)]
YEAR_OPTS   = ["All"] + [str(y) for y in range(2022, datetime.date.today().year + 3)]

# ══════════════════════════════════════════════════════════════════════════════
#  DATABASE
# ══════════════════════════════════════════════════════════════════════════════
def _db():
    return sqlite3.connect(DB_PATH)

def _row_factory(cursor, row):
    return dict(zip([d[0] for d in cursor.description], row))

def init_db():
    with _db() as c:
        c.execute("""
            CREATE TABLE IF NOT EXISTS issues (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                date_reported  TEXT DEFAULT '',
                system_number  TEXT DEFAULT '',
                product_family TEXT DEFAULT '',
                issue_type     TEXT DEFAULT '',
                issue_desc     TEXT DEFAULT '',
                sps_number     TEXT DEFAULT '',
                ncr_number     TEXT DEFAULT '',
                status         TEXT DEFAULT 'Open',
                solution       TEXT DEFAULT '',
                solution_date  TEXT DEFAULT '',
                crf            TEXT DEFAULT '',
                esw            TEXT DEFAULT '',
                scv            TEXT DEFAULT '',
                remarks        TEXT DEFAULT '',
                tracker_type   TEXT DEFAULT 'Engineering',
                created_at     TEXT DEFAULT (datetime('now','localtime'))
            )
        """)
        existing = {r[1] for r in c.execute("PRAGMA table_info(issues)")}
        for col, defn in [
            ("ncr_number",    "TEXT DEFAULT ''"),
            ("solution",      "TEXT DEFAULT ''"),
            ("solution_date", "TEXT DEFAULT ''"),
            ("crf",           "TEXT DEFAULT ''"),
            ("esw",           "TEXT DEFAULT ''"),
            ("scv",           "TEXT DEFAULT ''"),
            ("remarks",       "TEXT DEFAULT ''"),
            ("tracker_type",  "TEXT DEFAULT 'Engineering'"),
        ]:
            if col not in existing:
                c.execute(f"ALTER TABLE issues ADD COLUMN {col} {defn}")
        
        # Create index on tracker_type for performance
        c.execute("CREATE INDEX IF NOT EXISTS idx_tracker_type ON issues(tracker_type)")

def fetch_issues(status_f="All", family_f="All", itype_f="All",
                 month_f="All", year_f="All", tracker_type="Engineering"):
    q    = """SELECT id, date_reported, system_number, product_family, issue_type,
                     issue_desc, sps_number, ncr_number, status,
                     solution, solution_date, crf, esw, scv, remarks, tracker_type, created_at
              FROM issues WHERE 1=1"""
    args = []
    if status_f != "All": q += " AND status=?";         args.append(status_f)
    if family_f != "All": q += " AND product_family=?"; args.append(family_f)
    if itype_f  != "All": q += " AND issue_type=?";     args.append(itype_f)
    if month_f  != "All":
        q += " AND CAST(strftime('%m', date_reported) AS INTEGER)=?"
        args.append(datetime.datetime.strptime(month_f, "%B").month)
    if year_f   != "All":
        q += " AND strftime('%Y', date_reported)=?";    args.append(year_f)
    q += " AND tracker_type=?"
    args.append(tracker_type)
    q += " ORDER BY date_reported DESC, id DESC"
    with _db() as c:
        c.row_factory = _row_factory
        return c.execute(q, args).fetchall()

def fetch_by_id(db_id):
    with _db() as c:
        c.row_factory = _row_factory
        return c.execute("SELECT * FROM issues WHERE id=?", (db_id,)).fetchone()

def insert_issue(date_reported, system_number, product_family, issue_type,
                 issue_desc, sps_number, ncr_number, tracker_type="Engineering"):
    with _db() as c:
        cur = c.execute(
            """INSERT INTO issues
               (date_reported, system_number, product_family, issue_type,
                issue_desc, sps_number, ncr_number, tracker_type)
               VALUES (?,?,?,?,?,?,?,?)""",
            (date_reported, system_number, product_family, issue_type,
             issue_desc, sps_number, ncr_number, tracker_type))
        return cur.lastrowid

def update_issue(db_id, date_reported, system_number, product_family, issue_type,
                 issue_desc, sps_number, ncr_number, status,
                 solution, solution_date, crf, esw, scv, remarks, tracker_type):
    with _db() as c:
        c.execute("""
            UPDATE issues SET
                date_reported=?, system_number=?, product_family=?, issue_type=?,
                issue_desc=?, sps_number=?, ncr_number=?, status=?,
                solution=?, solution_date=?, crf=?, esw=?, scv=?, remarks=?, tracker_type=?
            WHERE id=?""",
            (date_reported, system_number, product_family, issue_type,
             issue_desc, sps_number, ncr_number, status,
             solution, solution_date, crf, esw, scv, remarks, tracker_type, db_id))

def delete_by_ids(ids):
    with _db() as c:
        c.executemany("DELETE FROM issues WHERE id=?", [(i,) for i in ids])

def get_counts(tracker_type="Engineering"):
    with _db() as c:
        total  = c.execute("SELECT COUNT(*) FROM issues WHERE tracker_type=?", (tracker_type,)).fetchone()[0]
        open_  = c.execute("SELECT COUNT(*) FROM issues WHERE status='Open' AND tracker_type=?", (tracker_type,)).fetchone()[0]
        closed = c.execute("SELECT COUNT(*) FROM issues WHERE status='Closed' AND tracker_type=?", (tracker_type,)).fetchone()[0]
        clarif = c.execute("SELECT COUNT(*) FROM issues WHERE status='Clarification' AND tracker_type=?", (tracker_type,)).fetchone()[0]
    return total, open_, closed, clarif


# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL EXPORT
# ══════════════════════════════════════════════════════════════════════════════
def _fmt_date(d):
    try:    return datetime.datetime.strptime(d, "%Y-%m-%d").strftime("%d %b %Y")
    except: return d or ""

def _xl_border():
    s = Side(style="thin", color="B0C8E0")
    return Border(left=s, right=s, top=s, bottom=s)

def _auto_row_height(text, col_width_chars, base_pt=13, min_pt=18, padding=4):
    if not text:
        return min_pt
    lines = sum(max(1, -(-len(p) // max(1, col_width_chars)))
                for p in str(text).splitlines())
    return max(min_pt, lines * base_pt + padding)

def export_excel(engineering_rows, material_rows, filepath, customer_mode=False, export_scope="Both"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Production Issues"

    if customer_mode:
        headers    = ["Date Reported", "System Number", "Product Family",
                      "Issue Type", "Issue Description", "SPS Number", "Remarks"]
        col_widths = [14, 22, 14, 18, 60, 13, 35]
        left_cols  = {4, 6}
    else:
        headers    = ["Date Reported", "System Number", "Product Family", "Issue Type",
                      "Issue Description", "SPS Number", "NCR Number", "Status",
                      "Solution", "Solution Date", "CRF", "ESW", "SCV", "Remarks"]
        col_widths = [14, 22, 14, 18, 55, 13, 13, 10, 45, 14, 12, 12, 12, 35]
        left_cols  = {4, 8, 13}

    hdr_fill = PatternFill("solid", fgColor="1B5DA8")
    hdr_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    title_font = Font(name="Calibri", bold=True, size=16)
    ctr = Alignment(horizontal="center", vertical="top", wrap_text=True)
    lft = Alignment(horizontal="left",   vertical="top", wrap_text=True)
    title_align = Alignment(horizontal="left", vertical="center")

    current_row = 1

    # Helper function to write a block
    def write_block(rows, title, start_row):
        if not rows:
            return start_row
        
        # Title banner
        cell = ws.cell(row=start_row, column=1, value=title)
        cell.font = title_font
        cell.fill = PatternFill("solid", fgColor="007EA5")
        cell.alignment = title_align
        
        # Blank row gap
        start_row += 1
        
        # Headers row
        for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
            cell = ws.cell(row=start_row, column=ci, value=h)
            cell.font = hdr_font
            cell.fill = hdr_fill
            cell.alignment = ctr
            cell.border = _xl_border()
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[start_row].height = 26
        start_row += 1
        
        # Data rows
        open_fill   = PatternFill("solid", fgColor="FDEDEC")
        closed_fill = PatternFill("solid", fgColor="EAFAF1")
        alt_fill    = PatternFill("solid", fgColor="EBF4FF")
        
        for row in rows:
            fill = (open_fill   if row["status"] == "Open"   else
                    closed_fill if row["status"] == "Closed" else alt_fill)
            values = (
                [_fmt_date(row["date_reported"]), row["system_number"],
                 row["product_family"], row["issue_type"], row["issue_desc"],
                 row["sps_number"], row["remarks"]]
                if customer_mode else
                [_fmt_date(row["date_reported"]), row["system_number"],
                 row["product_family"], row["issue_type"], row["issue_desc"],
                 row["sps_number"], row["ncr_number"], row["status"],
                 row["solution"], _fmt_date(row["solution_date"]),
                 row["crf"], row["esw"], row["scv"], row["remarks"]]
            )
            row_h = max(
                (_auto_row_height(val, cw)
                 for ci0, (val, cw) in enumerate(zip(values, col_widths))
                 if ci0 in left_cols),
                default=18
            )
            ws.row_dimensions[start_row].height = row_h
            for ci, (val, cw) in enumerate(zip(values, col_widths), 1):
                cell = ws.cell(row=start_row, column=ci, value=val)
                cell.fill = fill
                cell.border = _xl_border()
                cell.font = Font(name="Calibri", size=10)
                cell.alignment = lft if (ci - 1) in left_cols else ctr
            start_row += 1
        
        return start_row

    # Write based on export scope
    if export_scope == "Engineering":
        current_row = write_block(engineering_rows, "ENGINEERING ISSUES LOG", current_row)
    elif export_scope == "Material":
        current_row = write_block(material_rows, "MATERIAL ISSUES LOG", current_row)
    else:  # Both
        current_row = write_block(engineering_rows, "ENGINEERING ISSUES LOG", current_row)
        # Two blank row gaps
        current_row += 2
        current_row = write_block(material_rows, "MATERIAL ISSUES LOG", current_row)

    ws.freeze_panes = f"A{3 if export_scope != 'Both' else 3}"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{current_row - 1}"
    wb.save(filepath)


# ══════════════════════════════════════════════════════════════════════════════
#  CANVAS TABLE
# ══════════════════════════════════════════════════════════════════════════════
class CanvasTable:
    """Excel-like table using tkinter.Canvas with gridlines."""
    def __init__(self, parent, columns, on_double_click=None, on_select=None):
        self.parent = parent
        self.columns = list(columns)
        self.on_double_click = on_double_click
        self.on_select = on_select
        self._data = []
        self._displayed_rows = []
        self._selected_idx = None
        self._id_map = {}
        self._wrap_cols = {c[0] for c in columns if c[3]}

        self.bg = C["surface"]
        self.header_bg = C["header"] 
        self.header_fg = "white"
        self.row_bg = C["surface"]
        self.alt_row_bg = C["stripe"]
        self.selected_bg = C["sel"]
        self.grid_color = C["border"]
        self.text_color = C["text"]
        self.highlight_color = C["accent"]

        self.default_row_height = 32
        self.header_height = 32
        self.serial_width = 40
        self._col_widths = [self.serial_width] + [c[2] for c in columns]
        self._row_heights = []
        self._cell_texts = {}

        self._resize_col_idx = None
        self._resize_row_idx = None
        self._resize_start_x = 0
        self._resize_start_y = 0
        self._resize_start_width = 0
        self._resize_start_height = 0
        self._resize_line_id = None
        self._sort_col = None
        self._sort_reverse = False

        self._highlight_rect_id = None
        self._selected_highlight_row = None
        self._selected_highlight_col = None
        self._selected_col_idx = None
        self._manual_row_heights = {}

        self._build()

    def _build(self):
        self.frame = ctk.CTkFrame(self.parent, fg_color=self.bg, corner_radius=10)

        self.canvas = tk.Canvas(self.frame, bg=self.bg, highlightthickness=0)
        self.vsb = tk.Scrollbar(self.frame, orient="vertical", command=self.canvas.yview)
        self.hsb = tk.Scrollbar(self.frame, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        self.canvas.grid(row=0, column=0, sticky="nsew", padx=(6, 0), pady=6)
        self.vsb.grid(row=0, column=1, sticky="ns", pady=6, padx=(0, 4))
        self.hsb.grid(row=1, column=0, sticky="ew", padx=(6, 0), pady=(0, 4))
        self.frame.grid_rowconfigure(0, weight=1)
        self.frame.grid_columnconfigure(0, weight=1)

        self.canvas.bind("<Button-1>", self._on_click)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<Double-1>", self._on_double_click)
        self.canvas.bind("<Configure>", lambda e: self._draw())
        self.canvas.bind("<Motion>", self._on_motion)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Button-4>", self._on_mousewheel)
        self.canvas.bind("<Button-5>", self._on_mousewheel)
        self.frame.winfo_toplevel().bind_all("<Control-c>", self._on_copy)
        self.frame.winfo_toplevel().bind_all("<Control-C>", self._on_copy)

        self._draw()

    def _total_width(self):
        return sum(self._col_widths)

    def _col_x_positions(self):
        xs = [0]
        for w in self._col_widths:
            xs.append(xs[-1] + w)
        return xs

    def _row_y_positions(self):
        ys = [self.header_height]
        for h in self._row_heights:
            ys.append(ys[-1] + h)
        return ys

    def _measure_text_height(self, text, width, font_family, font_size):
        if not text:
            return 0
        item = self.canvas.create_text(0, 0, text=text, width=width,
                                       font=(font_family, font_size),
                                       anchor="nw")
        bbox = self.canvas.bbox(item)
        self.canvas.delete(item)
        if bbox:
            return bbox[3] - bbox[1]
        return 0

    def _compute_row_heights(self):
        self._row_heights = []
        self._cell_texts = {}
        self._cell_heights = {}
        min_h = self.default_row_height
        pad_v = 8

        if not hasattr(self, 'canvas') or not self.canvas.winfo_exists():
            return

        for row_idx, data_idx in enumerate(self._displayed_rows):
            if row_idx in self._manual_row_heights:
                row = self._data[data_idx]
                self._cell_texts[(row_idx, 0)] = str(row_idx + 1)
                for col_idx, (cid, _, w, wrap) in enumerate(self.columns, start=1):
                    text = row.get(cid, "")
                    self._cell_texts[(row_idx, col_idx)] = text
                self._row_heights.append(self._manual_row_heights[row_idx])
                continue

            row = self._data[data_idx]
            max_text_h = 0

            self._cell_texts[(row_idx, 0)] = str(row_idx + 1)
            serial_h = self._measure_text_height(str(row_idx + 1), self.serial_width - 12,
                                                  F["family"], F["size_sm"])
            max_text_h = max(max_text_h, serial_h)
            self._cell_heights[(row_idx, 0)] = serial_h

            for col_idx, (cid, _, w, wrap) in enumerate(self.columns, start=1):
                text = row.get(cid, "")
                self._cell_texts[(row_idx, col_idx)] = text
                actual_width = self._col_widths[col_idx]

                if wrap and text:
                    text_h = self._measure_text_height(text, actual_width - 12,
                                                       F["family"], F["size_sm"])
                elif text:
                    text_h = self._measure_text_height(text, actual_width,
                                                       F["family"], F["size_sm"])
                else:
                    text_h = 0

                self._cell_heights[(row_idx, col_idx)] = text_h
                max_text_h = max(max_text_h, text_h)

            row_h = max(min_h, max_text_h + pad_v)
            self._row_heights.append(row_h)

    def _draw(self):
        self.canvas.delete("all")

        self._compute_row_heights()

        total_w = self._total_width()
        ys = self._row_y_positions()
        total_h = ys[-1] if ys else self.header_height

        self.canvas.configure(scrollregion=(0, 0, total_w, total_h))

        xs = self._col_x_positions()

        y = 0
        headers = [("#", self.serial_width)] + [(c[1], self._col_widths[i+1]) for i, c in enumerate(self.columns)]
        for i, (hdr, w) in enumerate(headers):
            x1, x2 = xs[i], xs[i+1]
            self.canvas.create_rectangle(x1, y, x2, y + self.header_height,
                                         fill=self.header_bg, outline="", tags="header")
            display_hdr = hdr
            if i > 0 and self._sort_col == self.columns[i-1][0]:
                display_hdr = f"{hdr} {'▼' if self._sort_reverse else '▲'}"
            self.canvas.create_text((x1 + x2) // 2, y + self.header_height // 2,
                                    text=display_hdr, fill=self.header_fg,
                                    font=(F["family"], F["size_sm"], "bold"),
                                    tags="header")
            if i > 0:
                handle_x = x2 - 3
                self.canvas.create_line(handle_x, y + 4, handle_x, y + self.header_height - 4,
                                        fill=self.header_bg, width=6, tags=("resize_handle", f"col_resize_{i}"))

        self.canvas.create_line(0, self.header_height, total_w, self.header_height,
                                fill=self.grid_color, width=1)

        for row_idx, data_idx in enumerate(self._displayed_rows):
            row = self._data[data_idx]
            y1, y2 = ys[row_idx], ys[row_idx + 1]
            row_h = y2 - y1
            is_selected = (row_idx == self._selected_idx)
            bg = self.selected_bg if is_selected else (self.alt_row_bg if row_idx % 2 else self.row_bg)

            self.canvas.create_rectangle(0, y1, total_w, y2,
                                         fill=bg, outline="", tags=f"row_{row_idx}")

            x1, x2 = xs[0], xs[1]
            text = self._cell_texts.get((row_idx, 0), str(row_idx + 1))
            self.canvas.create_text((x1 + x2) // 2, (y1 + y2) // 2,
                                    text=text, fill=self.text_color,
                                    font=(F["family"], F["size_sm"]),
                                    tags=(f"row_{row_idx}", "cell"))

            left_align_cols = {"desc", "solution", "remarks"}
            for col_idx, (cid, _, w, wrap) in enumerate(self.columns, start=1):
                x1, x2 = xs[col_idx], xs[col_idx + 1]
                text = self._cell_texts.get((row_idx, col_idx), "")
                cell_h = self._cell_heights.get((row_idx, col_idx), 0)
                is_left_align = cid in left_align_cols

                status = row.get("status", "")
                if cid == "status":
                    fg = {"Open": C["open"], "Clarification": C["clarif"], "Closed": C["closed"]}.get(text, self.text_color)
                elif status == "Open":
                    fg = C["open"]
                elif status == "Clarification":
                    fg = C["clarif"]
                elif status == "Closed":
                    fg = C["closed"]
                else:
                    fg = self.text_color

                col_width = x2 - x1
                if is_left_align:
                    self.canvas.create_text(x1 + 6, (y1 + y2) // 2,
                                            text=text, fill=fg, anchor="w",
                                            width=col_width - 12 if wrap else None,
                                            font=(F["family"], F["size_sm"]),
                                            tags=(f"row_{row_idx}", f"cell_{row_idx}_{col_idx}", "cell"))
                else:
                    self.canvas.create_text((x1 + x2) // 2, (y1 + y2) // 2,
                                            text=text, fill=fg, anchor="center",
                                            font=(F["family"], F["size_sm"]),
                                            tags=(f"row_{row_idx}", f"cell_{row_idx}_{col_idx}", "cell"))

            handle_y = y2 - 2
            self.canvas.create_line(0, handle_y, total_w, handle_y,
                                    fill=bg, width=4, tags=("resize_handle", f"row_resize_{row_idx}"))

        for x in xs:
            self.canvas.create_line(x, 0, x, total_h, fill=self.grid_color, width=1)

        for y in ys:
            self.canvas.create_line(0, y, total_w, y, fill=self.grid_color, width=1)

        for i in range(len(self._displayed_rows)):
            self.canvas.tag_bind(f"row_{i}", "<Button-1>", lambda e, idx=i: self._select_row(idx))

        if (self._selected_highlight_row is not None and
            self._selected_highlight_col is not None and
            self._selected_highlight_row < len(self._displayed_rows)):
            row_idx = self._selected_highlight_row
            col_idx = self._selected_highlight_col
            if col_idx < len(xs):
                self._show_cell_highlight(
                    row_idx, col_idx,
                    xs[col_idx], ys[row_idx],
                    xs[col_idx + 1], ys[row_idx + 1]
                )

    def _on_click(self, event):
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)

        items = self.canvas.find_overlapping(x-2, y-2, x+2, y+2)
        for item in items:
            tags = self.canvas.gettags(item)
            for tag in tags:
                if tag.startswith("col_resize_"):
                    self._resize_col_idx = int(tag.split("_")[-1])
                    self._resize_start_x = x
                    self._resize_start_width = self._col_widths[self._resize_col_idx]
                    return
                elif tag.startswith("row_resize_"):
                    self._resize_row_idx = int(tag.split("_")[-1])
                    self._resize_start_y = y
                    self._resize_start_height = self._row_heights[self._resize_row_idx]
                    return

        if y < self.header_height:
            xs = self._col_x_positions()
            for i, (cid, hdr, w, _) in enumerate(self.columns):
                if xs[i + 1] <= x < xs[i + 2]:
                    if self.on_select:
                        self.on_select("sort", cid)
                    return
        else:
            self._handle_cell_click(x, y)

    def _on_motion(self, event):
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)

        items = self.canvas.find_overlapping(x-3, y-3, x+3, y+3)
        for item in items:
            tags = self.canvas.gettags(item)
            if "resize_handle" in tags:
                self.canvas.configure(cursor="sb_h_double_arrow" if "col_resize" in str(tags) else "sb_v_double_arrow")
                return
        self.canvas.configure(cursor="")

    def _on_drag(self, event):
        if self._resize_col_idx is not None:
            x = self.canvas.canvasx(event.x)
            delta = x - self._resize_start_x
            new_width = max(30, self._resize_start_width + delta)
            self._col_widths[self._resize_col_idx] = new_width
            col_idx = self._resize_col_idx - 1
            if 0 <= col_idx < len(self.columns):
                c = self.columns[col_idx]
                self.columns[col_idx] = (c[0], c[1], new_width, c[3])
            self._draw()
        elif self._resize_row_idx is not None:
            y = self.canvas.canvasy(event.y)
            delta = y - self._resize_start_y
            new_height = max(20, self._resize_start_height + delta)
            self._manual_row_heights[self._resize_row_idx] = new_height
            self._draw()

    def _on_release(self, event):
        self._resize_col_idx = None
        self._resize_row_idx = None
        self._resize_start_x = 0
        self._resize_start_y = 0
        self.canvas.configure(cursor="")

    def _handle_cell_click(self, x, y):
        xs = self._col_x_positions()
        ys = self._row_y_positions()

        col_idx = None
        for i in range(len(xs) - 1):
            if xs[i] <= x < xs[i + 1]:
                col_idx = i
                break

        row_idx = None
        for i in range(len(ys) - 1):
            if ys[i] <= y < ys[i + 1]:
                row_idx = i
                break

        if col_idx is not None and row_idx is not None:
            self._selected_col_idx = col_idx
            text = self._cell_texts.get((row_idx, col_idx), "")
            if text:
                self.canvas.clipboard_clear()
                self.canvas.clipboard_append(text)

            self._show_cell_highlight(row_idx, col_idx, xs[col_idx], ys[row_idx], xs[col_idx + 1], ys[row_idx + 1])

    def _show_cell_highlight(self, row_idx, col_idx, x1, y1, x2, y2):
        self._clear_cell_highlight()

        self._selected_highlight_row = row_idx
        self._selected_highlight_col = col_idx

        self._highlight_rect_id = self.canvas.create_rectangle(
            x1 + 1, y1 + 1, x2 - 1, y2 - 1,
            outline=self.highlight_color, width=2, fill="",
            tags="cell_highlight"
        )
        self.canvas.tag_raise("cell_highlight")
        for item in self.canvas.find_withtag("cell"):
            self.canvas.tag_raise(item)

    def _clear_cell_highlight(self):
        if self._highlight_rect_id:
            self.canvas.delete(self._highlight_rect_id)
            self._highlight_rect_id = None
        self._selected_highlight_row = None
        self._selected_highlight_col = None

    def _on_double_click(self, event):
        if self._selected_idx is not None and self.on_double_click:
            self.on_double_click(self._selected_idx)

    def _on_copy(self, event):
        self._clear_cell_highlight()

        if self._selected_idx is None:
            return

        col_idx = self._selected_col_idx if self._selected_col_idx is not None else 1

        text = self._cell_texts.get((self._selected_idx, col_idx), "")
        if text:
            self.canvas.clipboard_clear()
            self.canvas.clipboard_append(text)

    def _on_mousewheel(self, event):
        if hasattr(event, 'delta') and event.delta:
            scroll_units = int(-1 * (event.delta / 120))
        elif event.num == 4:
            scroll_units = -1
        elif event.num == 5:
            scroll_units = 1
        else:
            return
        self.canvas.yview_scroll(scroll_units, "units")

    def _select_row(self, row_idx):
        self._selected_idx = row_idx
        self._selected_col_idx = None
        self._selected_highlight_row = None
        self._selected_highlight_col = None
        self._draw()
        if self.on_select:
            db_id = self._id_map.get(row_idx)
            self.on_select("select", db_id)

    def clear(self):
        self._data = []
        self._displayed_rows = []
        self._selected_idx = None
        self._selected_col_idx = None
        self._selected_highlight_row = None
        self._selected_highlight_col = None
        self._id_map = {}
        self._manual_row_heights = {}
        self._draw()

    def add_row(self, db_id, data):
        idx = len(self._data)
        self._data.append(data)
        self._displayed_rows.append(idx)
        self._id_map[len(self._displayed_rows) - 1] = db_id
        self._draw()

    def get_db_id(self, displayed_idx):
        return self._id_map.get(displayed_idx)

    def get_selected_db_ids(self):
        if self._selected_idx is not None:
            db_id = self._id_map.get(self._selected_idx)
            return [db_id] if db_id else []
        return []

    def clear_selection(self):
        self._selected_idx = None
        self._draw()

    def pack(self, **kw):
        self.frame.pack(**kw)

    def filter_rows(self, matching_indices):
        self._displayed_rows = matching_indices
        self._selected_idx = None
        self._selected_highlight_row = None
        self._selected_highlight_col = None
        self._manual_row_heights = {}
        self._id_map = {i: self._data[idx]["_db_id"] for i, idx in enumerate(matching_indices) if "_db_id" in self._data[idx]}
        self._draw()

    def show_all_rows(self):
        self._displayed_rows = list(range(len(self._data)))
        self._selected_idx = None
        self._selected_highlight_row = None
        self._selected_highlight_col = None
        self._manual_row_heights = {}
        self._id_map = {i: self._data[idx]["_db_id"] for i, idx in enumerate(self._displayed_rows) if "_db_id" in self._data[idx]}
        self._draw()

    def sort_by(self, col_id, reverse=False):
        self._sort_col = col_id
        self._sort_reverse = reverse
        self._selected_highlight_row = None
        self._selected_highlight_col = None
        self._manual_row_heights = {}
        def key_fn(data_idx):
            return self._data[data_idx].get(col_id, "")
        self._displayed_rows.sort(key=key_fn, reverse=reverse)
        self._id_map = {i: self._data[idx]["_db_id"] for i, idx in enumerate(self._displayed_rows) if "_db_id" in self._data[idx]}
        self._draw()

    def get_row_count(self):
        return len(self._displayed_rows)


# ══════════════════════════════════════════════════════════════════════════════
#  UI PRIMITIVES
# ══════════════════════════════════════════════════════════════════════════════
def _font(**kw):
    return ctk.CTkFont(family=F["family"], **kw)

def _lbl(parent, text, bold=False, size=None, color=None, **kw):
    return ctk.CTkLabel(parent, text=text, anchor="w",
                        text_color=color or C["text"],
                        font=_font(size=size or F["size_md"],
                                   weight="bold" if bold else "normal"), **kw)

def _entry(parent, ph="", width=None, state="normal"):
    kw = dict(fg_color=C["entry_bg"], text_color=C["text"], border_color=C["border"],
              corner_radius=8, placeholder_text=ph, state=state,
              font=_font(size=F["size_md"]))
    if width:
        kw["width"] = width
    e = ctk.CTkEntry(parent, **kw)
    stack = []
    def _push(*_):
        v = e.get()
        if not stack or stack[-1] != v:
            stack.append(v)
            if len(stack) > 200: stack.pop(0)
    def _undo(_):
        if len(stack) > 1:
            stack.pop(); e.delete(0, "end"); e.insert(0, stack[-1])
        return "break"
    e.bind("<KeyRelease>", _push)
    e.bind("<Control-z>",  _undo)
    return e

def _textbox(parent, h=80):
    tb = ctk.CTkTextbox(parent, height=h, fg_color=C["entry_bg"],
                        text_color=C["text"], border_color=C["border"],
                        border_width=1, corner_radius=8,
                        font=_font(size=F["size_md"]),
                        wrap="word")
    tb._textbox.configure(undo=True, maxundo=-1)
    return tb

def _btn(parent, text, cmd, outline=False, width=120, color=None, radius=10, **kw):
    if outline:
        return ctk.CTkButton(parent, text=text, command=cmd, width=width,
                             corner_radius=radius, fg_color="transparent",
                             border_width=1, border_color=C["accent"],
                             text_color=C["accent"], hover_color=C["panel"],
                             font=_font(size=F["size_md"]), **kw)
    return ctk.CTkButton(parent, text=text, command=cmd, width=width,
                         corner_radius=radius, fg_color=color or C["accent"],
                         hover_color=C["accent_h"], text_color="white",
                         font=_font(size=F["size_md"]), **kw)

def _section_bar(parent, title, row):
    f = ctk.CTkFrame(parent, fg_color=C["header"], corner_radius=0, height=24)
    f.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(10, 2))
    _lbl(f, f"  {title}", size=F["size_sm"], bold=True, color="white"
         ).pack(side="left", padx=8, pady=2)

def _optmenu(parent, vals, var, width=None):
    return ToggleDropdown(parent, vals, var, width=width)

def _popup_xy(widget):
    widget.update_idletasks()
    return widget.winfo_rootx(), widget.winfo_rooty() + widget.winfo_height() + 4


# ══════════════════════════════════════════════════════════════════════════════
#  TOGGLE DROPDOWN
# ══════════════════════════════════════════════════════════════════════════════
class ToggleDropdown(ctk.CTkFrame):
    def __init__(self, master, values, variable, width=None):
        super().__init__(master, fg_color="transparent")
        self._values = values
        self._var    = variable
        self._popup  = None
        w = width or 150
        self._btn = ctk.CTkButton(
            self, text=self._label(), width=w, anchor="w",
            fg_color=C["accent"], hover_color=C["accent_h"],
            text_color="white", corner_radius=8,
            font=_font(size=F["size_md"]),
            command=self._toggle)
        self._btn.pack(fill="both", expand=True)
        variable.trace_add("write", lambda *_: self._btn.configure(text=self._label()))

    def _label(self):
        return f"  {self._var.get()}  ▾"

    def _toggle(self):
        if self._popup and self._popup.winfo_exists():
            self._close()
        else:
            self._show()

    def _show(self):
        x, y = _popup_xy(self._btn)
        bw   = self._btn.winfo_width()

        row_h     = 32
        max_rows  = 10
        n         = len(self._values)
        visible   = min(n, max_rows)
        pop_h     = visible * row_h

        self._popup = tk.Toplevel(self)
        self._popup.overrideredirect(True)
        self._popup.geometry(f"{bw}x{pop_h}+{x}+{y}")
        self._popup.configure(bg=C["border"])
        self._popup.lift()

        canvas = tk.Canvas(self._popup, bg=C["surface"], highlightthickness=0,
                           width=bw-2, height=pop_h)
        vsb    = tk.Scrollbar(self._popup, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)

        if n > max_rows:
            vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(canvas, bg=C["surface"])
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_inner_resize(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(win_id, width=canvas.winfo_width())
        inner.bind("<Configure>", _on_inner_resize)
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(win_id, width=e.width))

        def _wheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind("<MouseWheel>", _wheel)
        inner.bind("<MouseWheel>",  _wheel)

        for val in self._values:
            rf = tk.Frame(inner, bg=C["surface"], cursor="hand2", height=row_h)
            rf.pack(fill="x"); rf.pack_propagate(False)
            lb = tk.Label(rf, text=f"  {val}", bg=C["surface"], fg=C["text"],
                          font=(F["family"], F["size_sm"]), anchor="w")
            lb.pack(fill="both", expand=True)
            for w2 in (rf, lb):
                w2.bind("<Enter>",    lambda e, r=rf, l=lb:
                        (r.configure(bg=C["sel"]),     l.configure(bg=C["sel"])))
                w2.bind("<Leave>",    lambda e, r=rf, l=lb:
                        (r.configure(bg=C["surface"]), l.configure(bg=C["surface"])))
                w2.bind("<Button-1>", lambda e, v=val: self._pick(v))
                w2.bind("<MouseWheel>", _wheel)

        self._popup.after(50, lambda: self._popup.focus_force()
                          if self._popup and self._popup.winfo_exists() else None)
        self._popup.bind("<FocusOut>", self._on_focusout)

    def _on_focusout(self, _):
        if not (self._popup and self._popup.winfo_exists()): return
        bx, by = self._btn.winfo_rootx(), self._btn.winfo_rooty()
        bw, bh = self._btn.winfo_width(),  self._btn.winfo_height()
        mx, my = self._popup.winfo_pointerxy()
        self._close(skip_toggle=(bx <= mx <= bx+bw and by <= my <= by+bh))

    def _pick(self, v):
        self._var.set(v); self._close()

    def _close(self, skip_toggle=False):
        if self._popup and self._popup.winfo_exists():
            self._popup.destroy()
        self._popup = None
        if skip_toggle:
            self._btn.configure(state="disabled")
            self.after(150, lambda: self._btn.configure(state="normal"))


# ══════════════════════════════════════════════════════════════════════════════
#  CALENDAR WIDGET
# ══════════════════════════════════════════════════════════════════════════════
class CalendarPopup(tk.Toplevel):
    def __init__(self, parent, anchor_widget, initial=None):
        super().__init__(parent)
        self.overrideredirect(True)
        self.configure(bg=C["surface"])
        self._result = None
        self._view   = initial or datetime.date.today()
        self._sel    = initial

        self._build()
        self._render()

        self.update_idletasks()
        x, y = _popup_xy(anchor_widget)
        self.geometry(f"+{x}+{y}")
        self.lift()

        self.grab_set()
        self.bind("<Button-1>", self._on_click)
        self.bind("<Escape>",   lambda _: self._close())

    def _build(self):
        nav = tk.Frame(self, bg=C["header"])
        nav.pack(fill="x")
        for side, txt, cmd in [("left", "◀", self._prev), ("right", "▶", self._next)]:
            tk.Button(nav, text=txt, bg=C["header"], fg="white", relief="flat",
                      bd=0, padx=8, font=(F["family"], 11, "bold"),
                      activebackground=C["accent_h"], activeforeground="white",
                      command=cmd).pack(side=side, pady=5)
        self._hdr = tk.Label(nav, text="", bg=C["header"], fg="white",
                             font=(F["family"], 11, "bold"), width=18)
        self._hdr.pack(side="left", expand=True)

        dow = tk.Frame(self, bg=C["panel"])
        dow.pack(fill="x")
        for i, d in enumerate(["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]):
            tk.Label(dow, text=d, width=4, bg=C["panel"], fg=C["subtle"],
                     font=(F["family"], 9, "bold")).grid(row=0, column=i, padx=3, pady=3)

        self._grid = tk.Frame(self, bg=C["surface"])
        self._grid.pack(padx=6, pady=4)

        foot = tk.Frame(self, bg=C["surface"])
        foot.pack(fill="x", pady=(0, 6))
        tk.Button(foot, text="Today", bg=C["panel"], fg=C["text"], relief="flat",
                  font=(F["family"], 9), activebackground=C["border"],
                  command=self._pick_today).pack()

    def _render(self):
        for w in self._grid.winfo_children():
            w.destroy()
        y, m  = self._view.year, self._view.month
        today = datetime.date.today()
        self._hdr.config(text=f"{calendar.month_name[m]}  {y}")
        for r, week in enumerate(calendar.monthcalendar(y, m)):
            for cc, day in enumerate(week):
                if day == 0:
                    tk.Label(self._grid, text="", width=4, bg=C["surface"]
                             ).grid(row=r, column=cc, padx=2, pady=2)
                    continue
                d   = datetime.date(y, m, day)
                sel = d == self._sel
                tod = d == today
                bg  = C["accent"] if sel else (C["sel"] if tod else C["surface"])
                fg  = "white" if sel else C["text"]
                fw  = "bold"  if sel or tod else "normal"
                tk.Button(self._grid, text=str(day), width=4, bg=bg, fg=fg,
                          relief="flat", bd=0, font=(F["family"], 10, fw),
                          activebackground=C["accent"], activeforeground="white",
                          command=lambda dd=d: self._pick(dd)
                          ).grid(row=r, column=cc, padx=2, pady=2)

    def _on_click(self, event):
        wx, wy = self.winfo_rootx(), self.winfo_rooty()
        ww, wh = self.winfo_width(),  self.winfo_height()
        if not (wx <= event.x_root <= wx + ww and wy <= event.y_root <= wy + wh):
            self._close()

    def _pick(self, d):
        self._result = d.strftime("%Y-%m-%d")
        self._close()

    def _pick_today(self):
        self._result = datetime.date.today().strftime("%Y-%m-%d")
        self._close()

    def _close(self):
        self.grab_release()
        self.destroy()

    def _prev(self):
        y, m = self._view.year, self._view.month - 1
        if m == 0: m, y = 12, y - 1
        self._view = datetime.date(y, m, 1); self._render()

    def _next(self):
        y, m = self._view.year, self._view.month + 1
        if m == 13: m, y = 1, y + 1
        self._view = datetime.date(y, m, 1); self._render()

    def result(self):
        return self._result


# ══════════════════════════════════════════════════════════════════════════════
#  DATE FIELD
# ══════════════════════════════════════════════════════════════════════════════
class DateField(ctk.CTkFrame):
    def __init__(self, master, initial=None, nullable=False, **kw):
        super().__init__(master, fg_color="transparent", **kw)
        self._val      = initial or ("" if nullable else datetime.date.today().strftime("%Y-%m-%d"))
        self._var      = tk.StringVar(value=_fmt_date(self._val))
        self._nullable = nullable

        ctk.CTkEntry(self, textvariable=self._var, state="readonly", width=116,
                     fg_color=C["entry_bg"], text_color=C["text"],
                     border_color=C["border"], corner_radius=8).pack(side="left")
        self._cal_btn = ctk.CTkButton(self, text="📅", width=32, corner_radius=8,
                                      fg_color=C["accent"], hover_color=C["accent_h"],
                                      command=self._open)
        self._cal_btn.pack(side="left", padx=(4, 0))
        if nullable:
            ctk.CTkButton(self, text="✕", width=26, corner_radius=8,
                          fg_color=C["panel"], hover_color=C["border"],
                          text_color=C["text"],
                          command=self._clear).pack(side="left", padx=(2, 0))

    def _open(self):
        try:
            init = datetime.datetime.strptime(self._val, "%Y-%m-%d").date() if self._val else None
        except ValueError:
            init = None
        popup = CalendarPopup(self, anchor_widget=self._cal_btn, initial=init)
        self.wait_window(popup)
        if popup.result():
            self._val = popup.result()
            self._var.set(_fmt_date(self._val))

    def _clear(self):
        self._val = ""; self._var.set("")

    def get(self):    return self._val
    def set(self, v): self._val = v; self._var.set(v)


# ══════════════════════════════════════════════════════════════════════════════
#  COPY-PASTE TEXT GENERATOR
# ══════════════════════════════════════════════════════════════════════════════
class CopyPasteDialog(ctk.CTkToplevel):
    def __init__(self, parent, system_number, issue_desc, sps_number):
        super().__init__(parent)
        self.title("Copy-Paste Text")
        self.grab_set()
        self.configure(fg_color=C["bg"])
        self._build(system_number, issue_desc, sps_number)

    def _build(self, sys_num, desc, sps_num):
        systems_block = sys_num.strip()
        email_text = (
            f"Hi KB / Zach,\n\n"
            f"SPS {sps_num} submitted for the following issue(s):\n\n"
            f"{systems_block}\n\n"
            f"{desc}\n\n"
            f"Best regards,"
        )
        desc_of_def = f"{systems_block}\n\n{desc}"
        det_dispos  = f"AMAT SPS {sps_num} submitted."

        outer = ctk.CTkFrame(self, fg_color=C["bg"], corner_radius=0)
        outer.pack(fill="both", expand=True, padx=10, pady=10)

        def _block(title, content, h=120):
            _lbl(outer, title, bold=True, size=F["size_sm"], color=C["header"]
                 ).pack(anchor="w", padx=18, pady=(14, 2))
            tb = _textbox(outer, h=h)
            tb.pack(fill="x", padx=18)
            tb.insert("1.0", content)
            def _copy():
                self.clipboard_clear()
                self.clipboard_append(tb.get("1.0", "end").rstrip())
                btn.configure(text="Copied ✓")
                self.after(1500, lambda: btn.configure(text="Copy"))
            btn = _btn(outer, "Copy", _copy, outline=True, width=80, radius=7)
            btn.pack(anchor="e", padx=18, pady=(3, 0))

        _block("📧  Email body",                            email_text,  h=170)
        _block("📋  NCR — Desc. of Def. / Req. for Change", desc_of_def, h=140)
        _block("📋  NCR — Det. Dispos. / Reas. for Change", det_dispos,  h=60)

        self.update_idletasks()
        height = outer.winfo_reqheight() + 60
        self.geometry(f"640x{height}")
        self.resizable(False, False)


# ══════════════════════════════════════════════════════════════════════════════
#  NEW ISSUE DIALOG
# ══════════════════════════════════════════════════════════════════════════════
class NewIssueDialog(ctk.CTkToplevel):
    def __init__(self, parent, tracker_type="Engineering"):
        super().__init__(parent)
        self.tracker_type = tracker_type
        self.title(f"Log New {tracker_type} Issue")
        self.geometry("640x750")
        self.resizable(False, True)
        self.grab_set()
        self.configure(fg_color=C["bg"])
        self._build()

    def _build(self):
        self.grid_columnconfigure(1, weight=1)
        P = dict(padx=22, pady=4)
        r = 0

        _lbl(self, f"New {self.tracker_type} Issue", bold=True, size=F["size_xl"]).grid(
            row=r, column=0, columnspan=2, padx=22, pady=(18, 10), sticky="w"); r += 1

        _lbl(self, "Date Reported *").grid(row=r, column=0, **P, sticky="w")
        self.dt = DateField(self)
        self.dt.grid(row=r, column=1, **P, sticky="w"); r += 1

        _lbl(self, "System Number(s) *").grid(row=r, column=0, **P, sticky="nw")
        sf = ctk.CTkFrame(self, fg_color="transparent")
        sf.grid(row=r, column=1, **P, sticky="ew")
        self.sys = _textbox(sf, h=60); self.sys.pack(fill="x")
        _lbl(sf, "One system per line for multiple systems.",
             size=F["size_sm"]-1, color=C["subtle"]).pack(anchor="w", pady=(1, 0)); r += 1

        _lbl(self, "Product Family *").grid(row=r, column=0, **P, sticky="w")
        self.fam = ctk.StringVar(value=FAMILIES[0])
        _optmenu(self, FAMILIES, self.fam).grid(row=r, column=1, **P, sticky="ew"); r += 1

        _lbl(self, "Issue Type *").grid(row=r, column=0, **P, sticky="w")
        self.ityp = ctk.StringVar(value=ISSUE_TYPES[0])
        _optmenu(self, ISSUE_TYPES, self.ityp).grid(row=r, column=1, **P, sticky="ew"); r += 1

        _lbl(self, "Issue Description *").grid(row=r, column=0, **P, sticky="nw")
        df = ctk.CTkFrame(self, fg_color="transparent")
        df.grid(row=r, column=1, **P, sticky="ew")
        self.desc = _textbox(df, h=130); self.desc.pack(fill="x")
        _lbl(df, "Separate multiple issues with a blank line.",
             size=F["size_sm"]-1, color=C["subtle"]).pack(anchor="w", pady=(1, 0)); r += 1

        _lbl(self, "SPS Number *").grid(row=r, column=0, **P, sticky="w")
        self.sps = _entry(self, "e.g. 752766")
        self.sps.grid(row=r, column=1, **P, sticky="ew"); r += 1

        _lbl(self, "NCR Number").grid(row=r, column=0, **P, sticky="w")
        self.ncr = _entry(self, "e.g. NCR281174  (fill after creating in Agile)")
        self.ncr.grid(row=r, column=1, **P, sticky="ew"); r += 1

        bf = ctk.CTkFrame(self, fg_color="transparent")
        bf.grid(row=r, column=0, columnspan=2, pady=18)
        _btn(bf, "Cancel",                   self.destroy,    outline=True, width=90
             ).pack(side="left", padx=6)
        _btn(bf, "Generate Copy-Paste Text", self._generate,
             color=C["subtle"], width=210).pack(side="left", padx=6)
        _btn(bf, "Submit",                   self._submit,    width=120
             ).pack(side="left", padx=6)

    def _generate(self):
        CopyPasteDialog(self,
                        self.sys.get("1.0", "end").strip(),
                        self.desc.get("1.0", "end").strip(),
                        self.sps.get().strip())

    def _submit(self):
        d    = self.dt.get().strip()
        s    = self.sys.get("1.0", "end").strip()
        desc = self.desc.get("1.0", "end").strip()
        sps  = self.sps.get().strip()
        if not all([d, s, desc, sps]):
            messagebox.showerror("Missing Fields",
                                 "Date Reported, System Number, Issue Description "
                                 "and SPS Number are required.", parent=self)
            return
        insert_issue(d, s, self.fam.get(), self.ityp.get(), desc,
                     sps, self.ncr.get().strip(), self.tracker_type)
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
#  MANAGE / EDIT DIALOG
# ══════════════════════════════════════════════════════════════════════════════
class ManageDialog(ctk.CTkToplevel):
    def __init__(self, parent, db_id):
        super().__init__(parent)
        self.db_id = db_id
        self.title("Edit Issue")
        self.geometry("720x880")
        self.resizable(True, True)
        self.grab_set()
        self.configure(fg_color=C["bg"])
        self.R = fetch_by_id(db_id)
        self._build()

    def _build(self):
        sf = ctk.CTkScrollableFrame(self, fg_color=C["bg"], corner_radius=0,
                                    scrollbar_button_color=C["accent"],
                                    scrollbar_button_hover_color=C["accent_h"])
        sf.pack(fill="both", expand=True)
        sf.grid_columnconfigure(1, weight=1)
        P = dict(padx=22, pady=4)
        R = self.R; r = 0

        title_sys = R["system_number"].splitlines()[0] if R["system_number"] else ""
        tracker_type = R.get("tracker_type", "Engineering")
        _lbl(sf, f"{tracker_type} Issue — {title_sys}", bold=True, size=F["size_lg"]).grid(
            row=r, column=0, columnspan=2, padx=22, pady=(16, 2), sticky="w"); r += 1
        _lbl(sf, f"Created: {R['created_at']}", size=F["size_sm"], color=C["subtle"]).grid(
            row=r, column=0, columnspan=2, padx=22, sticky="w"); r += 1

        _section_bar(sf, "Basic Information", r); r += 1

        _lbl(sf, "Date Reported").grid(row=r, column=0, **P, sticky="w")
        self.dt = DateField(sf, initial=R["date_reported"])
        self.dt.grid(row=r, column=1, **P, sticky="w"); r += 1

        _lbl(sf, "System Number(s)").grid(row=r, column=0, **P, sticky="nw")
        sys_f = ctk.CTkFrame(sf, fg_color="transparent")
        sys_f.grid(row=r, column=1, **P, sticky="ew")
        self.sys = _textbox(sys_f, h=60); self.sys.pack(fill="x")
        self.sys.insert("1.0", R["system_number"])
        _lbl(sys_f, "One system per line for multiple systems.",
             size=F["size_sm"]-1, color=C["subtle"]).pack(anchor="w", pady=(1, 0)); r += 1

        _lbl(sf, "Product Family").grid(row=r, column=0, **P, sticky="w")
        self.fam = ctk.StringVar(value=R["product_family"])
        _optmenu(sf, FAMILIES, self.fam).grid(row=r, column=1, **P, sticky="ew"); r += 1

        _lbl(sf, "Issue Type").grid(row=r, column=0, **P, sticky="w")
        self.ityp = ctk.StringVar(value=R["issue_type"])
        _optmenu(sf, ISSUE_TYPES, self.ityp).grid(row=r, column=1, **P, sticky="ew"); r += 1

        _lbl(sf, "Issue Description").grid(row=r, column=0, **P, sticky="nw")
        df = ctk.CTkFrame(sf, fg_color="transparent")
        df.grid(row=r, column=1, **P, sticky="ew")
        self.desc = _textbox(df, h=100); self.desc.pack(fill="x")
        self.desc.insert("1.0", R["issue_desc"])
        _lbl(df, "Separate multiple issues with a blank line.",
             size=F["size_sm"]-1, color=C["subtle"]).pack(anchor="w", pady=(1, 0)); r += 1

        _lbl(sf, "SPS Number").grid(row=r, column=0, **P, sticky="w")
        self.sps = _entry(sf); self.sps.insert(0, R["sps_number"])
        self.sps.grid(row=r, column=1, **P, sticky="ew"); r += 1

        _lbl(sf, "NCR Number").grid(row=r, column=0, **P, sticky="w")
        self.ncr = _entry(sf); self.ncr.insert(0, R["ncr_number"])
        self.ncr.grid(row=r, column=1, **P, sticky="ew"); r += 1

        _lbl(sf, "Status").grid(row=r, column=0, **P, sticky="w")
        self.status = ctk.StringVar(value=R["status"])
        rf = ctk.CTkFrame(sf, fg_color="transparent")
        rf.grid(row=r, column=1, **P, sticky="w")
        for v, col in [("Open", C["open"]), ("Clarification", C["clarif"]), ("Closed", C["closed"])]:
            ctk.CTkRadioButton(rf, text=v, value=v, variable=self.status,
                               fg_color=col, hover_color=col, text_color=C["text"],
                               corner_radius=50, font=_font(size=F["size_md"])
                               ).pack(side="left", padx=10)
        r += 1

        _section_bar(sf, "Resolution", r); r += 1

        _lbl(sf, "Solution").grid(row=r, column=0, **P, sticky="nw")
        self.sol = _textbox(sf, h=80)
        self.sol.insert("1.0", R["solution"] or "")
        self.sol.grid(row=r, column=1, **P, sticky="ew"); r += 1

        _lbl(sf, "Solution Date").grid(row=r, column=0, **P, sticky="w")
        self.sol_dt = DateField(sf, initial=R["solution_date"] or None, nullable=True)
        self.sol_dt.grid(row=r, column=1, **P, sticky="w"); r += 1

        _section_bar(sf, "Change Controls", r); r += 1
        for attr, label, val in [("crf", "CRF", R["crf"]),
                                  ("esw", "ESW", R["esw"]),
                                  ("scv", "SCV", R["scv"])]:
            _lbl(sf, label).grid(row=r, column=0, **P, sticky="w")
            e = _entry(sf); e.insert(0, val or "")
            e.grid(row=r, column=1, **P, sticky="ew")
            setattr(self, attr, e); r += 1

        _section_bar(sf, "Remarks", r); r += 1
        _lbl(sf, "Remarks").grid(row=r, column=0, **P, sticky="nw")
        self.rmk = _textbox(sf, h=70)
        self.rmk.insert("1.0", R["remarks"] or "")
        self.rmk.grid(row=r, column=1, **P, sticky="ew"); r += 1

        bb = ctk.CTkFrame(self, fg_color=C["panel"], corner_radius=0, height=52)
        bb.pack(fill="x", side="bottom"); bb.pack_propagate(False)
        _btn(bb, "Delete",       self._delete, color="#a02020", width=100, radius=8
             ).pack(side="left",  padx=14, pady=10)
        _btn(bb, "Cancel",       self.destroy, outline=True,    width=90,  radius=8
             ).pack(side="right", padx=8,  pady=10)
        _btn(bb, "Save Changes", self._save,                    width=150, radius=8
             ).pack(side="right", padx=8,  pady=10)

    def _save(self):
        tracker_type = self.R.get("tracker_type", "Engineering")
        update_issue(self.db_id,
            self.dt.get(),
            self.sys.get("1.0", "end").strip(),
            self.fam.get(), self.ityp.get(),
            self.desc.get("1.0", "end").strip(),
            self.sps.get().strip(), self.ncr.get().strip(),
            self.status.get(),
            self.sol.get("1.0", "end").strip(),
            self.sol_dt.get(),
            self.crf.get().strip(), self.esw.get().strip(), self.scv.get().strip(),
            self.rmk.get("1.0", "end").strip(),
            tracker_type)
        messagebox.showinfo("Saved", "Issue updated.", parent=self)
        self.destroy()

    def _delete(self):
        if messagebox.askyesno("Delete", "Permanently delete this issue?", parent=self):
            delete_by_ids([self.db_id]); self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
#  TAB CONTENT FRAME
# ══════════════════════════════════════════════════════════════════════════════
class TabContentFrame(ctk.CTkFrame):
    """Container for a single tab's filters, stats, and table."""
    def __init__(self, parent, tracker_type, on_double_click, on_table_event, on_new_issue, on_mass_delete):
        super().__init__(parent, fg_color="transparent")
        self.tracker_type = tracker_type
        self.on_double_click = on_double_click
        self.on_table_event = on_table_event
        self.on_new_issue = on_new_issue
        self.on_mass_delete = on_mass_delete
        self._build()

    def _build(self):
        # Stats bar
        self._build_stats()
        
        # Filters
        self._build_filters()
        
        # Toolbar
        self._build_toolbar()
        
        # Table
        self._build_table()

    def _build_stats(self):
        bar = tk.Frame(self, bg=C["panel"], height=34)
        bar.pack(fill="x"); bar.pack_propagate(False)
        self._v_total  = tk.StringVar(value="Total: 0")
        self._v_open   = tk.StringVar(value="Open: 0")
        self._v_clarif = tk.StringVar(value="Clarification: 0")
        self._v_closed = tk.StringVar(value="Closed: 0")
        for var, col in [(self._v_total,  C["text"]),
                         (self._v_open,   C["open"]),
                         (self._v_clarif, C["clarif"]),
                         (self._v_closed, C["closed"])]:
            tk.Label(bar, textvariable=var, bg=C["panel"], fg=col,
                     font=(F["family"], F["size_md"], "bold")
                     ).pack(side="left", padx=18, pady=6)

    def _build_filters(self):
        bar = ctk.CTkFrame(self, fg_color=C["surface"], corner_radius=0)
        bar.pack(fill="x", ipady=6)
        row = ctk.CTkFrame(bar, fg_color="transparent")
        row.pack(pady=(12,0), padx=14)
        _lbl(row, "Filters:", bold=True, size=F["size_sm"], color=C["subtle"]
             ).grid(row=0, column=0, padx=(0, 10))

        def flt(label, var, opts, col, w=None):
            _lbl(row, label, size=F["size_sm"], color=C["subtle"]
                 ).grid(row=0, column=col, padx=(0, 3))
            _optmenu(row, opts, var, width=w).grid(row=0, column=col+1, padx=(0, 12))
            var.trace_add("write", lambda *_: self._refresh())

        self._fs = ctk.StringVar(value="All")
        self._ff = ctk.StringVar(value="All")
        self._fi = ctk.StringVar(value="All")
        self._fm = ctk.StringVar(value="All")
        self._fy = ctk.StringVar(value="All")
        flt("Status", self._fs, ["All", "Open", "Clarification", "Closed"],  1, 100)
        flt("Family", self._ff, ["All"] + FAMILIES,         3, 100)
        flt("Issue",  self._fi, ["All"] + ISSUE_TYPES,      5, 175)
        flt("Month",  self._fm, MONTH_NAMES,                7, 120)
        flt("Year",   self._fy, YEAR_OPTS,                  9,  88)
        _btn(row, "Reset", self._reset_filters, outline=True, width=68, radius=8
             ).grid(row=0, column=11, padx=(4, 0))

    def _build_toolbar(self):
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", padx=14, pady=(10, 0))

        _btn(bar, "+ New Issue", lambda: self.on_new_issue(self.tracker_type),
             width=180, radius=8).pack(side="left")
        _btn(bar, "Delete Selected", lambda: self.on_mass_delete(self.tracker_type),
             color="#a02020", width=148, radius=8).pack(side="left", padx=(10, 0))
        self._sel_lbl = _lbl(bar, "", size=F["size_sm"], color=C["subtle"])
        self._sel_lbl.pack(side="left", padx=10)

        _lbl(bar, "Click row to select, double-click to edit",
             size=F["size_sm"], color=C["subtle"]).pack(side="right", padx=(0, 16))

        self._search_match_lbl = _lbl(bar, "", size=F["size_sm"], color=C["subtle"])
        self._search_match_lbl.pack(side="right", padx=(0, 6))

        self._search_clear_btn = _btn(bar, "✕", self._search_clear,
                                       outline=True, width=28, radius=6)
        self._search_clear_btn.pack(side="right", padx=(0, 2))
        self._search_clear_btn.pack_forget()

        self._search_var = tk.StringVar()
        self._search_entry = ctk.CTkEntry(
            bar, textvariable=self._search_var,
            placeholder_text="🔍  Search all columns…",
            width=240, corner_radius=8,
            fg_color=C["entry_bg"], text_color=C["text"],
            border_color=C["border"],
            font=_font(size=F["size_md"]))
        self._search_entry.pack(side="right", padx=(0, 4))
        self._search_var.trace_add("write", lambda *_: self._search_apply())

        self._search_entry.bind("<Escape>", lambda e: self._search_clear())

    def _build_table(self):
        outer = ctk.CTkFrame(self, fg_color=C["surface"], corner_radius=10)
        outer.pack(fill="both", expand=True, padx=14, pady=8)

        self._wrap_cols = {c[0] for c in TABLE_COLS if c[3]}
        self._col_hdrs  = {c[0]: c[1] for c in TABLE_COLS}
        self._sort_col = None
        self._sort_rev = False

        self.table = CanvasTable(outer, TABLE_COLS,
                                  on_double_click=self.on_double_click,
                                  on_select=self.on_table_event)
        self.table.pack(fill="both", expand=True, padx=6, pady=6)

    def _refresh(self, *_):
        self.table.clear()
        self._search_var.set("")
        self._search_match_lbl.configure(text="")
        self._search_clear_btn.pack_forget()
        self._sort_col = None
        self._sort_rev = False

        rows = fetch_issues(self._fs.get(), self._ff.get(),
                           self._fi.get(), self._fm.get(), self._fy.get(),
                           self.tracker_type)
        for row in rows:
            data = {
                "_db_id"  : row["id"],
                "date"    : _fmt_date(row["date_reported"]),
                "system"  : row["system_number"]  or "",
                "family"  : row["product_family"],
                "type"    : row["issue_type"],
                "sps"     : row["sps_number"],
                "ncr"     : row["ncr_number"],
                "status"  : row["status"],
                "desc"    : row["issue_desc"]      or "",
                "solution": row["solution"]         or "",
                "sol_date": _fmt_date(row["solution_date"]),
                "crf"     : row["crf"],
                "esw"     : row["esw"],
                "scv"     : row["scv"],
                "remarks" : row["remarks"]          or "",
            }
            self.table.add_row(row["id"], data)

        t, o, cl, cf = get_counts(self.tracker_type)
        self._v_total.set(f"Total: {t}")
        self._v_open.set(f"Open: {o}")
        self._v_clarif.set(f"Clarification: {cf}")
        self._v_closed.set(f"Closed: {cl}")
        self._sel_lbl.configure(text="")

    def _reset_filters(self):
        for v in [self._fs, self._ff, self._fi, self._fm, self._fy]:
            v.set("All")

    def _search_apply(self, *_):
        query = self._search_var.get().strip().lower()

        if query:
            self._search_clear_btn.pack(side="right", padx=(0, 2),
                                         before=self._search_entry)
        else:
            self._search_clear_btn.pack_forget()

        if not query:
            self.table.show_all_rows()
            n = self.table.get_row_count()
            self._search_match_lbl.configure(text="")
            return

        tokens = query.split()
        matches = []

        for i, row in enumerate(self.table._data):
            haystack = " ".join(str(v) for v in row.values()).lower()
            if all(t in haystack for t in tokens):
                matches.append(i)

        self.table.filter_rows(matches)

        n = len(matches)
        self._search_match_lbl.configure(
            text=f"{n} match{'es' if n != 1 else ''}" if n else "No matches",
            text_color=C["subtle"] if n else C["open"])

    def _search_clear(self):
        self._search_var.set("")
        self._search_entry.focus_set()
        self.table.show_all_rows()
        self._search_match_lbl.configure(text="")

    def get_selected_db_ids(self):
        return self.table.get_selected_db_ids()

    def update_selection_label(self, n):
        self._sel_lbl.configure(
            text=f"{n} row{'s' if n != 1 else ''} selected" if n else "")


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN WINDOW
# ══════════════════════════════════════════════════════════════════════════════
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.configure(fg_color=C["bg"])
        self.title(WIN_TITLE)
        self.geometry(WIN_SIZE)
        self.minsize(*WIN_MIN)
        self._id_map   = {}
        self._raw_text = {}
        self._tab_frames = {}
        self._build()
        self._refresh_all_tabs()

    def _build(self):
        self._build_header()
        self._build_tabview()

    def _build_header(self):
        bar = tk.Frame(self, bg=C["header"], height=54)
        bar.pack(fill="x"); bar.pack_propagate(False)

        if os.path.exists(LOGO_PATH):
            try:
                from PIL import ImageTk
                img = Image.open(LOGO_PATH)
                img.thumbnail((96, 72), Image.LANCZOS)
                self._logo = ImageTk.PhotoImage(img)
                tk.Label(bar, image=self._logo, bg=C["header"],
                         borderwidth=0).pack(side="left", padx=(14, 6), pady=9)
            except Exception:
                pass

        tk.Label(bar, text=WIN_TITLE, bg=C["header"], fg="white",
                 font=(F["family"], F["size_xl"], "bold")
                 ).pack(side="left", padx=(0, 20), pady=(20,0))

        for txt, cmd, bold in [
            ("Refresh All",  self._refresh_all_tabs,   False),
            ("Export Excel", self._export,             False),
        ]:
            tk.Button(bar, text=txt,
                      bg="#174e8a" if bold else C["header"],
                      fg="white", relief="flat", bd=0, padx=14, pady=5,
                      activebackground=C["accent_h"], activeforeground="white",
                      font=(F["family"], F["size_md"], "bold" if bold else "normal"),
                      command=cmd).pack(side="right", padx=(0, 8), pady=12)

    def _build_tabview(self):
        # Create container frame with grey background
        container = ctk.CTkFrame(self, fg_color=C["surface"], corner_radius=10)
        container.pack(fill="both", expand=True, padx=14, pady=8)
        
        # Create tabview inside container
        self.tabview = ctk.CTkTabview(container)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=15)
        
        # Add tabs
        self.tabview.add("Engineering Issues")
        self.tabview.add("Material Issues")
        
        # Create content frames for each tab
        for tab_name, tracker_type in [("Engineering Issues", "Engineering"), ("Material Issues", "Material")]:
            tab_frame = TabContentFrame(
                self.tabview.tab(tab_name),
                tracker_type,
                self._on_table_double_click,
                self._on_table_event,
                self._new_issue,
                self._mass_delete
            )
            tab_frame.pack(fill="both", expand=True)
            self._tab_frames[tracker_type] = tab_frame

    def _on_table_double_click(self, row_idx):
        # Get the active tab's tracker type
        active_tab = self.tabview.get()
        tracker_type = "Engineering" if active_tab == "Engineering Issues" else "Material"
        tab_frame = self._tab_frames[tracker_type]
        db_id = tab_frame.table.get_db_id(row_idx)
        if db_id:
            self._open_by_id(db_id)

    def _on_table_event(self, event_type, data):
        if event_type == "sort":
            # Get the active tab's tracker type
            active_tab = self.tabview.get()
            tracker_type = "Engineering" if active_tab == "Engineering Issues" else "Material"
            tab_frame = self._tab_frames[tracker_type]
            self._sort(tab_frame, data)
        elif event_type == "select":
            # Get the active tab's tracker type
            active_tab = self.tabview.get()
            tracker_type = "Engineering" if active_tab == "Engineering Issues" else "Material"
            tab_frame = self._tab_frames[tracker_type]
            n = 1 if data else 0
            tab_frame.update_selection_label(n)

    def _sort(self, tab_frame, col):
        rev = (not tab_frame._sort_rev) if tab_frame._sort_col == col else False
        tab_frame.table.sort_by(col, reverse=rev)
        tab_frame._sort_col = col
        tab_frame._sort_rev = rev

    def _refresh_all_tabs(self, *_):
        for tracker_type, tab_frame in self._tab_frames.items():
            tab_frame._refresh()

    def _new_issue(self, tracker_type):
        dlg = NewIssueDialog(self, tracker_type)
        self.wait_window(dlg)
        self._tab_frames[tracker_type]._refresh()

    def _open_by_id(self, db_id):
        dlg = ManageDialog(self, db_id)
        self.wait_window(dlg)
        # Refresh the appropriate tab based on the issue's tracker type
        issue = fetch_by_id(db_id)
        if issue:
            tracker_type = issue.get("tracker_type", "Engineering")
            if tracker_type in self._tab_frames:
                self._tab_frames[tracker_type]._refresh()

    def _mass_delete(self, tracker_type):
        tab_frame = self._tab_frames[tracker_type]
        ids = tab_frame.get_selected_db_ids()
        if not ids:
            messagebox.showinfo("No Selection", "Select rows to delete first.",
                                parent=self)
            return
        if messagebox.askyesno("Confirm Delete",
                               f"Permanently delete {len(ids)} {tracker_type} issue(s)?\n"
                               "This cannot be undone.", parent=self):
            delete_by_ids(ids)
            tab_frame._refresh()

    def _export(self):
        # Show export scope selection dialog
        export_dialog = ExportScopeDialog(self)
        self.wait_window(export_dialog)
        
        if not export_dialog.result:
            return
        
        export_scope = export_dialog.result
        
        # Fetch data for both types
        engineering_rows = fetch_issues(tracker_type="Engineering") if export_scope in ["Engineering", "Both"] else []
        material_rows = fetch_issues(tracker_type="Material") if export_scope in ["Material", "Both"] else []
        
        # Determine if customer mode
        customer_mode = False
        active_tab = self.tabview.get()
        tracker_type = "Engineering" if active_tab == "Engineering Issues" else "Material"
        tab_frame = self._tab_frames[tracker_type]
        
        if tab_frame._fs.get() == "Open":
            answer = messagebox.askyesnocancel(
                "Export Type",
                "You are exporting Open issues.\n\n"
                "Export as Customer Report?\n"
                "(Yes = customer view, no internal fields)\n"
                "(No  = full internal log)",
                parent=self)
            if answer is None: return
            customer_mode = answer
        
        suffix = "_Customer" if customer_mode else ""
        scope_suffix = f"_{export_scope}" if export_scope != "Both" else ""
        path = filedialog.asksaveasfilename(
            parent=self, defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile=f"AMAT_Issues{scope_suffix}{suffix}_{datetime.date.today()}.xlsx")
        if not path: return
        
        try:
            export_excel(engineering_rows, material_rows, path, customer_mode=customer_mode, export_scope=export_scope)
            total_rows = len(engineering_rows) + len(material_rows)
            if messagebox.askyesno("Exported",
                                   f"Saved {total_rows} rows.\n\nOpen file now?",
                                   parent=self):
                os.startfile(path)
        except Exception as e:
            messagebox.showerror("Export Error", str(e), parent=self)


# ══════════════════════════════════════════════════════════════════════════════
#  EXPORT SCOPE DIALOG
# ══════════════════════════════════════════════════════════════════════════════
class ExportScopeDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Select Export Scope")
        self.geometry("400x250")
        self.resizable(False, False)
        self.grab_set()
        self.configure(fg_color=C["bg"])
        self.result = None
        self._build()

    def _build(self):
        outer = ctk.CTkFrame(self, fg_color=C["bg"], corner_radius=0)
        outer.pack(fill="both", expand=True, padx=20, pady=20)

        _lbl(outer, "Select Export Scope", bold=True, size=F["size_xl"]).pack(
            anchor="w", padx=10, pady=(10, 20))

        _lbl(outer, "Choose which issue types to include in the export:",
             size=F["size_md"], color=C["subtle"]).pack(anchor="w", padx=10, pady=(0, 20))

        self._scope_var = tk.StringVar(value="Both")

        options = [
            ("Engineering Issues Only", "Engineering"),
            ("Material Issues Only", "Material"),
            ("Both Engineering and Material Issues", "Both")
        ]

        for text, value in options:
            rb = ctk.CTkRadioButton(
                outer, text=text, value=value, variable=self._scope_var,
                fg_color=C["accent"], hover_color=C["accent_h"], text_color=C["text"],
                corner_radius=50, font=_font(size=F["size_md"])
            )
            rb.pack(anchor="w", padx=10, pady=8)

        bf = ctk.CTkFrame(outer, fg_color="transparent")
        bf.pack(fill="x", pady=(20, 0))

        _btn(bf, "Cancel", self._cancel, outline=True, width=90, radius=8
             ).pack(side="right", padx=6)
        _btn(bf, "Export", self._export, width=120, radius=8
             ).pack(side="right", padx=6)

    def _export(self):
        self.result = self._scope_var.get()
        self.destroy()

    def _cancel(self):
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    init_db()
    App().mainloop()
