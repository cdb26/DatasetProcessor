import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
import shutil
import threading
import queue
from collections import defaultdict

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    _XLSX_OK = True
except ImportError:
    _XLSX_OK = False

# ─── Theme ───────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

ACCENT   = "#4F8EF7"
ACCENT2  = "#7C5CFC"
BG_DARK  = "#0F1117"
BG_MID   = "#161B27"
BG_CARD  = "#1E2535"
BG_FIELD = "#252D3D"
TEXT_MAIN = "#E8EDF5"
TEXT_DIM  = "#7A8599"
SUCCESS   = "#3DD68C"
WARNING   = "#F5A623"
DANGER    = "#F75C5C"
BORDER    = "#2E3A50"

# ─── Filename parser ──────────────────────────────────────────────────────────
FILENAME_RE = re.compile(
    r"^(\d{6})_([\d.]+m)_(\d)_(close|medium|far)_(dim|well)_(\d{4})(_depth)?(\.(?:jpg|png))$",
    re.IGNORECASE,
)

def parse_filename(name: str) -> dict | None:
    m = FILENAME_RE.match(name)
    if not m:
        return None
    return {
        "room":     m.group(1),
        "height":   m.group(2),
        "angle":    m.group(3),
        "distance": m.group(4),
        "lighting": m.group(5),
        "sequence": m.group(6),
        "is_depth": m.group(7) is not None,   # True when "_depth" suffix present
        "ext":      m.group(8).lower(),
        "original": name,
    }

def build_filename(parts: dict) -> str:
    depth_suffix = "_depth" if parts.get("is_depth") else ""
    return (f"{parts['room']}_{parts['height']}_{parts['angle']}_"
            f"{parts['distance']}_{parts['lighting']}_{parts['sequence']}"
            f"{depth_suffix}{parts['ext']}")

ANGLE_LABEL  = {"1": "Ortho", "2": "Diagonal", "3": "Top-down"}
DIST_LABEL   = {"close": "Close", "medium": "Medium", "far": "Far"}
LIGHT_LABEL  = {"dim": "Dim", "well": "Well-lit"}
HEIGHT_OPTS  = ["0.8m", "1.2m", "1.6m"]
ANGLE_OPTS   = ["1", "2", "3"]
DIST_OPTS    = ["close", "medium", "far"]
LIGHT_OPTS   = ["dim", "well"]


def walk_images(root: str):
    for dirpath, _, files in os.walk(root):
        for f in sorted(files):
            if f.lower().endswith((".jpg", ".png")):
                yield os.path.join(dirpath, f), os.path.relpath(dirpath, root), f


# ═══════════════════════════════════════════════════════════════════════════════
class DatasetManagerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Dataset Manager")
        self.geometry("1240x820")
        self.minsize(980, 680)
        self.configure(fg_color=BG_DARK)
        self.dataset_path = tk.StringVar(value="")
        self.status_var   = tk.StringVar(value="Ready")
        self._build_ui()

    # ── Layout ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        hdr = ctk.CTkFrame(self, fg_color=BG_MID, corner_radius=0, height=60)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="⬡  DATASET MANAGER",
                     font=("Courier New", 17, "bold"),
                     text_color=ACCENT).pack(side="left", padx=24)
        ctk.CTkLabel(hdr, text="Image filename toolkit for structured data collection",
                     font=("Courier New", 11), text_color=TEXT_DIM).pack(side="left", padx=4)

        path_row = ctk.CTkFrame(self, fg_color=BG_MID, corner_radius=0, height=52)
        path_row.pack(fill="x")
        path_row.pack_propagate(False)
        ctk.CTkLabel(path_row, text="Dataset root:", font=("Courier New", 11),
                     text_color=TEXT_DIM).pack(side="left", padx=(20, 6), pady=14)
        ctk.CTkEntry(path_row, textvariable=self.dataset_path,
                     font=("Courier New", 11), fg_color=BG_FIELD,
                     text_color=TEXT_MAIN, border_color=BORDER,
                     width=560).pack(side="left", pady=14)
        ctk.CTkButton(path_row, text="Browse", width=90,
                      fg_color=ACCENT, hover_color=ACCENT2,
                      font=("Courier New", 11, "bold"),
                      command=self._browse).pack(side="left", padx=8)

        self.tabs = ctk.CTkTabview(
            self, fg_color=BG_CARD,
            segmented_button_fg_color=BG_MID,
            segmented_button_selected_color=ACCENT,
            segmented_button_unselected_color=BG_MID,
            segmented_button_selected_hover_color=ACCENT2,
            text_color=TEXT_MAIN, corner_radius=8)
        self.tabs.pack(fill="both", expand=True, padx=16, pady=(10, 0))
        self.tabs.add("✦  Rename")
        self.tabs.add("⊞  Filter / Search")
        self.tabs.add("⇄  Move by Room")
        self.tabs.add("◈  Preview")

        self._build_rename_tab(self.tabs.tab("✦  Rename"))
        self._build_filter_tab(self.tabs.tab("⊞  Filter / Search"))
        self._build_move_tab(self.tabs.tab("⇄  Move by Room"))
        self._build_preview_tab(self.tabs.tab("◈  Preview"))

        sb = ctk.CTkFrame(self, fg_color=BG_MID, corner_radius=0, height=30)
        sb.pack(fill="x", side="bottom")
        sb.pack_propagate(False)
        ctk.CTkLabel(sb, textvariable=self.status_var,
                     font=("Courier New", 10), text_color=TEXT_DIM).pack(side="left", padx=16)

    def _browse(self):
        d = filedialog.askdirectory(title="Select dataset root folder")
        if d:
            self.dataset_path.set(d)
            self._set_status(f"Dataset root: {d}")

    def _set_status(self, msg: str):
        self.status_var.set(msg)

    # ══════════════════════════════════════════════════════════════════════════
    #  TAB 1 — RENAME
    # ══════════════════════════════════════════════════════════════════════════
    def _build_rename_tab(self, parent):
        parent.configure(fg_color=BG_CARD)
        left = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=0)
        left.pack(side="left", fill="both", expand=True, padx=(12, 6), pady=12)
        right = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=0, width=280)
        right.pack(side="right", fill="y", padx=(0, 12), pady=12)
        right.pack_propagate(False)

        self._section_label(left, "CURRENT  →  NEW  FIELD  VALUES")
        grid = ctk.CTkFrame(left, fg_color=BG_FIELD, corner_radius=8)
        grid.pack(fill="x", pady=(6, 0))

        fields = [
            ("Floor + Room  (FFRRRRR)", "new_room",   "e.g. 070202"),
            ("Height",                  "new_height",  HEIGHT_OPTS),
            ("Angle",                   "new_angle",   ANGLE_OPTS),
            ("Distance",                "new_dist",    DIST_OPTS),
            ("Lighting",                "new_light",   LIGHT_OPTS),
        ]
        self._rename_vars: dict[str, tk.StringVar] = {}
        for i, (label, key, opts) in enumerate(fields):
            ctk.CTkLabel(grid, text=label, font=("Courier New", 10),
                         text_color=TEXT_DIM, anchor="w").grid(
                             row=i, column=0, sticky="w", padx=14, pady=6)
            var = tk.StringVar(value="")
            self._rename_vars[key] = var
            if isinstance(opts, list):
                mb = ctk.CTkOptionMenu(grid, variable=var,
                                       values=["(keep)"] + opts,
                                       fg_color=BG_CARD, button_color=ACCENT,
                                       button_hover_color=ACCENT2,
                                       text_color=TEXT_MAIN,
                                       font=("Courier New", 11), width=180)
                mb.set("(keep)")
                mb.grid(row=i, column=1, sticky="w", padx=10, pady=6)
            else:
                ctk.CTkEntry(grid, textvariable=var,
                             font=("Courier New", 11), fg_color=BG_CARD,
                             text_color=TEXT_MAIN, border_color=BORDER,
                             placeholder_text=opts, width=180).grid(
                                 row=i, column=1, sticky="w", padx=10, pady=6)

        # ── Sequence ─── All / Selected toggle ────────────────────────────────
        seq_outer = ctk.CTkFrame(left, fg_color="transparent")
        seq_outer.pack(fill="x", pady=(10, 0))
        self._section_label(seq_outer, "SEQUENCE")
        seq_card = ctk.CTkFrame(seq_outer, fg_color=BG_FIELD, corner_radius=8)
        seq_card.pack(fill="x")

        self._seq_mode = tk.StringVar(value="all")
        radio_row = ctk.CTkFrame(seq_card, fg_color="transparent")
        radio_row.pack(anchor="w", padx=14, pady=(8, 4))
        ctk.CTkLabel(radio_row, text="Apply to:", font=("Courier New", 10),
                     text_color=TEXT_DIM).pack(side="left", padx=(0, 10))
        ctk.CTkRadioButton(radio_row, text="All sequences",
                           variable=self._seq_mode, value="all",
                           font=("Courier New", 11), text_color=TEXT_MAIN,
                           fg_color=ACCENT, hover_color=ACCENT2,
                           command=self._toggle_seq_fields).pack(side="left", padx=(0, 18))
        ctk.CTkRadioButton(radio_row, text="Selected range",
                           variable=self._seq_mode, value="selected",
                           font=("Courier New", 11), text_color=TEXT_MAIN,
                           fg_color=ACCENT2, hover_color=ACCENT,
                           command=self._toggle_seq_fields).pack(side="left")

        self._seq_range_frame = ctk.CTkFrame(seq_card, fg_color="transparent")
        self._seq_range_frame.pack(anchor="w", padx=14, pady=(0, 10))
        self._seq_start = tk.StringVar()
        self._seq_end   = tk.StringVar()
        for lbl, var, ph in [("From", self._seq_start, "e.g. 0617"),
                              ("To  ", self._seq_end,   "e.g. 0703")]:
            ctk.CTkLabel(self._seq_range_frame, text=lbl, font=("Courier New", 10),
                         text_color=TEXT_DIM).pack(side="left", padx=(0, 4))
            ctk.CTkEntry(self._seq_range_frame, textvariable=var, width=100,
                         font=("Courier New", 11), fg_color=BG_CARD,
                         text_color=TEXT_MAIN, border_color=BORDER,
                         placeholder_text=ph).pack(side="left", padx=(0, 14))
        self._toggle_seq_fields()

        # ── Filter rows to rename ─────────────────────────────────────────────
        filter_row = ctk.CTkFrame(left, fg_color="transparent")
        filter_row.pack(fill="x", pady=(10, 0))
        self._section_label(filter_row, "ONLY RENAME FILES MATCHING  (leave blank = all)")
        fi = ctk.CTkFrame(filter_row, fg_color=BG_FIELD, corner_radius=8)
        fi.pack(fill="x")
        filter_fields = [
            ("Room",     "rf_room",   "e.g. 070701"),
            ("Height",   "rf_height", HEIGHT_OPTS),
            ("Angle",    "rf_angle",  ANGLE_OPTS),
            ("Distance", "rf_dist",   DIST_OPTS),
            ("Lighting", "rf_light",  LIGHT_OPTS),
        ]
        self._rfilter_vars: dict[str, tk.StringVar] = {}
        for j, (label, key, opts) in enumerate(filter_fields):
            ctk.CTkLabel(fi, text=label, font=("Courier New", 10),
                         text_color=TEXT_DIM).grid(
                             row=j//3, column=(j%3)*2, sticky="w", padx=10, pady=6)
            var2 = tk.StringVar(value="")
            self._rfilter_vars[key] = var2
            if isinstance(opts, list):
                mb2 = ctk.CTkOptionMenu(fi, variable=var2,
                                        values=["(any)"] + opts,
                                        fg_color=BG_CARD, button_color=ACCENT2,
                                        button_hover_color=ACCENT,
                                        text_color=TEXT_MAIN,
                                        font=("Courier New", 11), width=130)
                mb2.set("(any)")
                mb2.grid(row=j//3, column=(j%3)*2+1, sticky="w", padx=6, pady=6)
            else:
                ctk.CTkEntry(fi, textvariable=var2, width=100,
                             font=("Courier New", 11), fg_color=BG_CARD,
                             text_color=TEXT_MAIN, border_color=BORDER,
                             placeholder_text=opts).grid(
                                 row=j//3, column=(j%3)*2+1, sticky="w", padx=6, pady=6)

        # ── Right panel options ───────────────────────────────────────────────
        self._section_label(right, "OPTIONS")
        opts_frame = ctk.CTkFrame(right, fg_color=BG_FIELD, corner_radius=8)
        opts_frame.pack(fill="x", pady=(4, 0))
        self._dry_run   = tk.BooleanVar(value=True)
        self._backup    = tk.BooleanVar(value=True)
        self._both_exts = tk.BooleanVar(value=True)
        for text, var in [("Dry run (preview only)", self._dry_run),
                           ("Backup originals",       self._backup),
                           ("Rename .jpg + .png pairs", self._both_exts)]:
            ctk.CTkCheckBox(opts_frame, text=text, variable=var,
                            font=("Courier New", 11), text_color=TEXT_MAIN,
                            fg_color=ACCENT, hover_color=ACCENT2).pack(
                                anchor="w", padx=14, pady=7)

        ctk.CTkButton(right, text="⟳  Preview Changes", height=38,
                      fg_color=BG_FIELD, border_width=1, border_color=ACCENT,
                      hover_color=BG_MID, font=("Courier New", 12, "bold"),
                      text_color=ACCENT,
                      command=self._preview_rename).pack(fill="x", pady=(18, 6))
        ctk.CTkButton(right, text="✔  Apply Rename", height=44,
                      fg_color=ACCENT, hover_color=ACCENT2,
                      font=("Courier New", 13, "bold"), text_color="white",
                      command=self._apply_rename).pack(fill="x", pady=6)

        self._section_label(left, "OPERATION LOG")
        self._rename_log = ctk.CTkTextbox(left, height=160,
                                          font=("Courier New", 10),
                                          fg_color=BG_FIELD, text_color=TEXT_MAIN,
                                          border_color=BORDER, corner_radius=6)
        self._rename_log.pack(fill="both", expand=True, pady=(4, 0))

    def _toggle_seq_fields(self):
        state = "normal" if self._seq_mode.get() == "selected" else "disabled"
        for child in self._seq_range_frame.winfo_children():
            child.configure(state=state)

    # ══════════════════════════════════════════════════════════════════════════
    #  TAB 2 — FILTER / SEARCH
    # ══════════════════════════════════════════════════════════════════════════
    def _build_filter_tab(self, parent):
        parent.configure(fg_color=BG_CARD)
        top = ctk.CTkFrame(parent, fg_color=BG_CARD)
        top.pack(fill="x", padx=12, pady=12)

        self._section_label(top, "FILTER  CRITERIA  (leave field as '(any)' to skip)")
        crit = ctk.CTkFrame(top, fg_color=BG_FIELD, corner_radius=8)
        crit.pack(fill="x")

        filter_defs = [
            ("Room",     "f_room",   None),
            ("Height",   "f_height", HEIGHT_OPTS),
            ("Angle",    "f_angle",  ANGLE_OPTS),
            ("Distance", "f_dist",   DIST_OPTS),
            ("Lighting", "f_light",  LIGHT_OPTS),
        ]
        self._filter_vars: dict[str, tk.StringVar] = {}
        for idx, (label, key, opts) in enumerate(filter_defs):
            ctk.CTkLabel(crit, text=label, font=("Courier New", 10),
                         text_color=TEXT_DIM, anchor="w").grid(
                             row=0, column=idx, sticky="w", padx=10, pady=(8, 0))
            var = tk.StringVar(value="")
            self._filter_vars[key] = var
            if opts:
                mb = ctk.CTkOptionMenu(crit, variable=var,
                                       values=["(any)"] + opts,
                                       fg_color=BG_CARD, button_color=ACCENT2,
                                       button_hover_color=ACCENT,
                                       text_color=TEXT_MAIN,
                                       font=("Courier New", 11), width=130)
                mb.set("(any)")
                mb.grid(row=1, column=idx, sticky="w", padx=10, pady=(0, 8))
            else:
                ctk.CTkEntry(crit, textvariable=var, width=110,
                             font=("Courier New", 11), fg_color=BG_CARD,
                             text_color=TEXT_MAIN, border_color=BORDER,
                             placeholder_text="—").grid(
                                 row=1, column=idx, sticky="w", padx=10, pady=(0, 8))

        # ── Sequence — All / Selected ─────────────────────────────────────────
        seq_f = ctk.CTkFrame(top, fg_color="transparent")
        seq_f.pack(fill="x", pady=(6, 0))
        self._section_label(seq_f, "SEQUENCE")
        seq_card = ctk.CTkFrame(seq_f, fg_color=BG_FIELD, corner_radius=8)
        seq_card.pack(fill="x")

        self._fseq_mode = tk.StringVar(value="all")
        rr = ctk.CTkFrame(seq_card, fg_color="transparent")
        rr.pack(anchor="w", padx=14, pady=(8, 4))
        ctk.CTkLabel(rr, text="Apply to:", font=("Courier New", 10),
                     text_color=TEXT_DIM).pack(side="left", padx=(0, 10))
        ctk.CTkRadioButton(rr, text="All sequences",
                           variable=self._fseq_mode, value="all",
                           font=("Courier New", 11), text_color=TEXT_MAIN,
                           fg_color=ACCENT, hover_color=ACCENT2,
                           command=self._toggle_fseq_fields).pack(side="left", padx=(0, 18))
        ctk.CTkRadioButton(rr, text="Selected range",
                           variable=self._fseq_mode, value="selected",
                           font=("Courier New", 11), text_color=TEXT_MAIN,
                           fg_color=ACCENT2, hover_color=ACCENT,
                           command=self._toggle_fseq_fields).pack(side="left")

        self._fseq_range_frame = ctk.CTkFrame(seq_card, fg_color="transparent")
        self._fseq_range_frame.pack(anchor="w", padx=14, pady=(0, 10))
        self._fseq_start = tk.StringVar()
        self._fseq_end   = tk.StringVar()
        for lbl, var, ph in [("From", self._fseq_start, "e.g. 0617"),
                              ("To  ", self._fseq_end,   "e.g. 0703")]:
            ctk.CTkLabel(self._fseq_range_frame, text=lbl, font=("Courier New", 10),
                         text_color=TEXT_DIM).pack(side="left", padx=(0, 4))
            ctk.CTkEntry(self._fseq_range_frame, textvariable=var, width=100,
                         font=("Courier New", 11), fg_color=BG_CARD,
                         text_color=TEXT_MAIN, border_color=BORDER,
                         placeholder_text=ph).pack(side="left", padx=(0, 14))
        self._toggle_fseq_fields()

        # Extension filter
        self._filter_ext = tk.StringVar(value="both")
        ext_row = ctk.CTkFrame(top, fg_color="transparent")
        ext_row.pack(fill="x", pady=(8, 0))
        ctk.CTkLabel(ext_row, text="Type:", font=("Courier New", 10),
                     text_color=TEXT_DIM).pack(side="left", padx=(0, 8))
        for val, txt in [("both", "Color + Depth"),
                          ("jpg",  "Color only (.jpg)"),
                          ("png",  "Depth only (.png)")]:
            ctk.CTkRadioButton(ext_row, text=txt, variable=self._filter_ext,
                               value=val, font=("Courier New", 11), text_color=TEXT_MAIN,
                               fg_color=ACCENT, hover_color=ACCENT2).pack(side="left", padx=12)

        btn_row = ctk.CTkFrame(top, fg_color="transparent")
        btn_row.pack(fill="x", pady=(10, 0))
        ctk.CTkButton(btn_row, text="⊞  Search", height=38, width=140,
                      fg_color=ACCENT, hover_color=ACCENT2,
                      font=("Courier New", 12, "bold"),
                      command=self._run_filter).pack(side="left")
        self._filter_count = tk.StringVar(value="")
        ctk.CTkLabel(btn_row, textvariable=self._filter_count,
                     font=("Courier New", 11), text_color=SUCCESS).pack(side="left", padx=16)
        ctk.CTkButton(btn_row, text="⊡  Export list", height=38, width=130,
                      fg_color=BG_FIELD, border_width=1, border_color=ACCENT,
                      hover_color=BG_MID, font=("Courier New", 11), text_color=ACCENT,
                      command=self._export_filter_list).pack(side="left", padx=8)
        ctk.CTkButton(btn_row, text="⊠  Copy matched", height=38, width=150,
                      fg_color=BG_FIELD, border_width=1, border_color=ACCENT2,
                      hover_color=BG_MID, font=("Courier New", 11), text_color=ACCENT2,
                      command=self._copy_matched).pack(side="left", padx=4)

        hint = ctk.CTkFrame(parent, fg_color="transparent")
        hint.pack(fill="x", padx=12, pady=(4, 0))
        self._section_label(hint, "RESULTS  —  click any row to open image")
        ctk.CTkLabel(hint, text="(color .jpg  /  depth .png)",
                     font=("Courier New", 9), text_color=TEXT_DIM).pack(
                         side="right", padx=4, pady=(8, 2))

        col_hdr = ctk.CTkFrame(parent, fg_color=BG_MID, corner_radius=0, height=26)
        col_hdr.pack(fill="x", padx=12)
        col_hdr.pack_propagate(False)
        for txt, w in [("Base filename (no ext)", 44), ("Room", 8), ("Height", 7),
                        ("Angle", 10), ("Distance", 9), ("Lighting", 8),
                        ("Seq", 6), ("Has", 7)]:
            ctk.CTkLabel(col_hdr, text=txt.upper(), font=("Courier New", 9, "bold"),
                         text_color=ACCENT, width=w*7, anchor="w").pack(side="left", padx=4)

        self._filter_scroll = ctk.CTkScrollableFrame(
            parent, fg_color=BG_FIELD, corner_radius=0)
        self._filter_scroll.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        # _filter_matches now stores deduplicated records:
        # list of dict with keys: base_key, parts, color_path, depth_path
        self._filter_matches: list[dict] = []
        self._filter_row_widgets: list[ctk.CTkFrame] = []

    def _toggle_fseq_fields(self):
        state = "normal" if self._fseq_mode.get() == "selected" else "disabled"
        for child in self._fseq_range_frame.winfo_children():
            child.configure(state=state)

    # ══════════════════════════════════════════════════════════════════════════
    #  TAB 3 — MOVE BY ROOM
    # ══════════════════════════════════════════════════════════════════════════
    def _build_move_tab(self, parent):
        parent.configure(fg_color=BG_CARD)
        left = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=0)
        left.pack(side="left", fill="both", expand=True, padx=(12, 6), pady=12)
        right = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=0, width=300)
        right.pack(side="right", fill="y", padx=(0, 12), pady=12)
        right.pack_propagate(False)

        # ── Room selector ─────────────────────────────────────────────────────
        self._section_label(left, "SELECT  ROOM(S)  TO  MOVE  (FFRRRR)")
        room_ctrl = ctk.CTkFrame(left, fg_color=BG_FIELD, corner_radius=8)
        room_ctrl.pack(fill="x", pady=(6, 0))

        rc1 = ctk.CTkFrame(room_ctrl, fg_color="transparent")
        rc1.pack(fill="x", padx=14, pady=8)
        ctk.CTkLabel(rc1, text="Room code (FFRRRR):", font=("Courier New", 10),
                     text_color=TEXT_DIM).pack(side="left", padx=(0, 10))
        self._move_room_entry = ctk.CTkEntry(rc1, width=140,
                                             font=("Courier New", 12, "bold"),
                                             fg_color=BG_CARD, text_color=TEXT_MAIN,
                                             border_color=ACCENT,
                                             placeholder_text="e.g. 070701")
        self._move_room_entry.pack(side="left", padx=(0, 10))
        ctk.CTkButton(rc1, text="+ Add", width=80,
                      fg_color=ACCENT2, hover_color=ACCENT,
                      font=("Courier New", 11, "bold"),
                      command=self._add_move_room).pack(side="left", padx=4)
        ctk.CTkButton(rc1, text="Scan rooms", width=110,
                      fg_color=BG_CARD, border_width=1, border_color=ACCENT,
                      hover_color=BG_MID, font=("Courier New", 11), text_color=ACCENT,
                      command=self._scan_rooms).pack(side="left", padx=4)

        self._section_label(left, "ROOMS  QUEUED  FOR  MOVE")
        self._rooms_listbox = ctk.CTkTextbox(left, height=100,
                                             font=("Courier New", 11),
                                             fg_color=BG_FIELD, text_color=TEXT_MAIN,
                                             border_color=BORDER, corner_radius=6)
        self._rooms_listbox.pack(fill="x", pady=(4, 0))
        self._rooms_listbox.configure(state="disabled")
        self._move_rooms: list[str] = []

        ctk.CTkButton(left, text="✕  Clear all rooms", width=140,
                      fg_color=BG_FIELD, border_width=1, border_color=DANGER,
                      hover_color=BG_MID, font=("Courier New", 11), text_color=DANGER,
                      command=self._clear_move_rooms).pack(anchor="w", pady=(4, 0))

        # ── Destination structure ─────────────────────────────────────────────
        self._section_label(left, "DESTINATION  STRUCTURE")
        dest_card = ctk.CTkFrame(left, fg_color=BG_FIELD, corner_radius=8)
        dest_card.pack(fill="x", pady=(4, 0))

        self._move_struct = tk.StringVar(value="room_folder")
        for val, txt, desc in [
            ("room_folder", "One folder per room",
             "dest/070701/color/…  &  dest/070701/depth_raw/…"),
            ("flat",        "Flat — all files together",
             "dest/070701_0.8m_1_close_dim_0617.jpg"),
            ("mirror",      "Mirror original structure",
             "Keeps relative sub-folders as-is"),
        ]:
            rb_row = ctk.CTkFrame(dest_card, fg_color="transparent")
            rb_row.pack(anchor="w", padx=14, pady=4)
            ctk.CTkRadioButton(rb_row, text=txt, variable=self._move_struct,
                               value=val, font=("Courier New", 11), text_color=TEXT_MAIN,
                               fg_color=ACCENT, hover_color=ACCENT2).pack(side="left")
            ctk.CTkLabel(rb_row, text=f"  ↳ {desc}", font=("Courier New", 9),
                         text_color=TEXT_DIM).pack(side="left", padx=6)

        opts_card = ctk.CTkFrame(left, fg_color=BG_FIELD, corner_radius=8)
        opts_card.pack(fill="x", pady=(10, 0))
        self._move_copy = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(opts_card, text="Copy instead of move  (keep originals in place)",
                        variable=self._move_copy,
                        font=("Courier New", 11), text_color=TEXT_MAIN,
                        fg_color=ACCENT, hover_color=ACCENT2).pack(
                            anchor="w", padx=14, pady=10)

        # ── Right: destination + actions ──────────────────────────────────────
        self._section_label(right, "DESTINATION  FOLDER")
        self._move_dest = tk.StringVar(value="")
        ctk.CTkEntry(right, textvariable=self._move_dest,
                     font=("Courier New", 10), fg_color=BG_FIELD,
                     text_color=TEXT_MAIN, border_color=BORDER,
                     placeholder_text="Click Browse…").pack(fill="x", pady=(4, 6))
        ctk.CTkButton(right, text="Browse destination…", height=36,
                      fg_color=BG_FIELD, border_width=1, border_color=ACCENT,
                      hover_color=BG_MID, font=("Courier New", 11), text_color=ACCENT,
                      command=self._browse_move_dest).pack(fill="x", pady=4)

        ctk.CTkButton(right, text="⟳  Preview Move", height=38,
                      fg_color=BG_FIELD, border_width=1, border_color=ACCENT2,
                      hover_color=BG_MID, font=("Courier New", 12, "bold"),
                      text_color=ACCENT2,
                      command=self._preview_move).pack(fill="x", pady=(20, 6))
        ctk.CTkButton(right, text="⇄  Execute Move", height=44,
                      fg_color=ACCENT2, hover_color=ACCENT,
                      font=("Courier New", 13, "bold"), text_color="white",
                      command=self._execute_move).pack(fill="x", pady=6)

        self._section_label(left, "MOVE  LOG")
        self._move_log = ctk.CTkTextbox(left, height=160,
                                        font=("Courier New", 10),
                                        fg_color=BG_FIELD, text_color=TEXT_MAIN,
                                        border_color=BORDER, corner_radius=6)
        self._move_log.pack(fill="both", expand=True, pady=(4, 0))

    def _add_move_room(self):
        code = self._move_room_entry.get().strip()
        if not re.fullmatch(r"\d{6}", code):
            messagebox.showerror("Invalid", "Room code must be exactly 6 digits (FFRRRR).")
            return
        if code in self._move_rooms:
            messagebox.showinfo("Duplicate", f"{code} is already in the list.")
            return
        self._move_rooms.append(code)
        self._refresh_rooms_list()
        self._move_room_entry.delete(0, "end")

    def _scan_rooms(self):
        root = self._get_root()
        if not root:
            return
        found: set[str] = set()
        for _, _, fname in walk_images(root):
            p = parse_filename(fname)
            if p:
                found.add(p["room"])
        if not found:
            messagebox.showinfo("None found", "No recognisable images in dataset root.")
            return

        win = ctk.CTkToplevel(self)
        win.title("Select rooms to add")
        win.geometry("340x420")
        win.configure(fg_color=BG_DARK)
        win.grab_set()
        ctk.CTkLabel(win, text="Rooms found in dataset:",
                     font=("Courier New", 11, "bold"), text_color=ACCENT).pack(pady=(14, 6))
        checks: dict[str, tk.BooleanVar] = {}
        scroll = ctk.CTkScrollableFrame(win, fg_color=BG_FIELD, corner_radius=8)
        scroll.pack(fill="both", expand=True, padx=16, pady=4)
        for code in sorted(found):
            v = tk.BooleanVar(value=(code not in self._move_rooms))
            checks[code] = v
            ctk.CTkCheckBox(scroll, text=code, variable=v,
                            font=("Courier New", 11), text_color=TEXT_MAIN,
                            fg_color=ACCENT, hover_color=ACCENT2).pack(anchor="w", pady=3)

        def _confirm():
            for code, v in checks.items():
                if v.get() and code not in self._move_rooms:
                    self._move_rooms.append(code)
            self._refresh_rooms_list()
            win.destroy()

        ctk.CTkButton(win, text="Add selected", fg_color=ACCENT, hover_color=ACCENT2,
                      font=("Courier New", 11, "bold"), command=_confirm).pack(
                          fill="x", padx=16, pady=10)

    def _clear_move_rooms(self):
        self._move_rooms.clear()
        self._refresh_rooms_list()

    def _refresh_rooms_list(self):
        self._rooms_listbox.configure(state="normal")
        self._rooms_listbox.delete("1.0", "end")
        for r in self._move_rooms:
            self._rooms_listbox.insert("end", f"  {r}\n")
        self._rooms_listbox.configure(state="disabled")

    def _browse_move_dest(self):
        d = filedialog.askdirectory(title="Select destination folder")
        if d:
            self._move_dest.set(d)

    def _gather_move_plan(self) -> list[tuple[str, str]] | None:
        root = self._get_root()
        if not root:
            return None
        dest = self._move_dest.get().strip()
        if not dest:
            messagebox.showerror("Error", "Please select a destination folder.")
            return None
        if not self._move_rooms:
            messagebox.showerror("Error", "No rooms selected. Add at least one room code.")
            return None

        struct = self._move_struct.get()
        plan: list[tuple[str, str]] = []

        for abs_path, rel_folder, fname in walk_images(root):
            parts = parse_filename(fname)
            if not parts:
                # File doesn't match naming convention — skip
                continue

            if parts["room"] not in self._move_rooms:
                continue

            if struct == "room_folder":
                # Always route by actual file extension, never by which subfolder it came from
                if fname.lower().endswith(".jpg"):
                    sub = "color"
                else:
                    sub = "depth_raw"
                dst = os.path.join(dest, parts["room"], sub, fname)
            elif struct == "flat":
                dst = os.path.join(dest, fname)
            else:  # mirror — preserve the exact relative folder from dataset root
                dst = os.path.join(dest, rel_folder, fname)

            plan.append((abs_path, dst))

        return plan

    def _preview_move(self):
        plan = self._gather_move_plan()
        if plan is None:
            return
        self._clear_log(self._move_log)
        if not plan:
            self._log(self._move_log, "No matching files found for selected rooms.")
            return
        verb = "COPY" if self._move_copy.get() else "MOVE"
        # count per type
        jpg_count = sum(1 for s, _ in plan if s.lower().endswith(".jpg"))
        png_count = sum(1 for s, _ in plan if s.lower().endswith(".png"))
        self._log(self._move_log,
                  f"Found {len(plan)} file(s): {jpg_count} color (.jpg)  +  {png_count} depth (.png)")
        self._log(self._move_log, f"{'SOURCE':<55}  →  DESTINATION ({verb})")
        self._log(self._move_log, "─" * 130)
        for src, dst in plan:
            tag = "[color]" if src.lower().endswith(".jpg") else "[depth]"
            self._log(self._move_log, f"  {tag}  {os.path.basename(src):<48}  →  {dst}")
        self._log(self._move_log,
                  f"\n{len(plan)} file(s) would be {'copied' if self._move_copy.get() else 'moved'}.")
        self._set_status(f"Move preview: {len(plan)} file(s)  ({jpg_count} color, {png_count} depth)")

    def _execute_move(self):
        plan = self._gather_move_plan()
        if plan is None:
            return
        if not plan:
            messagebox.showinfo("Nothing to do", "No matching files found.")
            return
        verb = "copy" if self._move_copy.get() else "move"
        jpg_count = sum(1 for s, _ in plan if s.lower().endswith(".jpg"))
        png_count = sum(1 for s, _ in plan if s.lower().endswith(".png"))
        if not messagebox.askyesno("Confirm",
                                   f"{verb.capitalize()} {len(plan)} file(s)?\n"
                                   f"  • {jpg_count} color (.jpg)\n"
                                   f"  • {png_count} depth (.png)\n\n"
                                   f"Rooms: {', '.join(self._move_rooms)}"):
            return
        self._clear_log(self._move_log)
        done = errors = 0
        for src, dst in plan:
            try:
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                if self._move_copy.get():
                    shutil.copy2(src, dst)
                else:
                    shutil.move(src, dst)
                tag = "[color]" if src.lower().endswith(".jpg") else "[depth]"
                self._log(self._move_log, f"✔  {tag}  {os.path.basename(src)}")
                done += 1
            except Exception as e:
                self._log(self._move_log,
                          f"✘  {os.path.basename(src)}  →  {dst}\n     ERROR: {e}")
                errors += 1
        summary = f"Done: {done} {verb}d, {errors} error(s)."
        self._log(self._move_log, "\n" + summary)
        self._set_status(summary)
        messagebox.showinfo("Complete", summary)

    # ══════════════════════════════════════════════════════════════════════════
    #  TAB 4 — PREVIEW
    # ══════════════════════════════════════════════════════════════════════════
    def _build_preview_tab(self, parent):
        parent.configure(fg_color=BG_CARD)
        ctrl = ctk.CTkFrame(parent, fg_color=BG_CARD)
        ctrl.pack(fill="x", padx=12, pady=12)
        ctk.CTkButton(ctrl, text="◈  Scan & Preview All Files", height=38, width=220,
                      fg_color=ACCENT, hover_color=ACCENT2,
                      font=("Courier New", 12, "bold"),
                      command=self._scan_all).pack(side="left")
        self._scan_count = tk.StringVar(value="")
        ctk.CTkLabel(ctrl, textvariable=self._scan_count,
                     font=("Courier New", 11), text_color=SUCCESS).pack(side="left", padx=16)

        hdr = ctk.CTkFrame(parent, fg_color=BG_MID, corner_radius=0, height=24)
        hdr.pack(fill="x", padx=12)
        hdr.pack_propagate(False)
        for txt, w in [("Path", 35), ("Room", 8), ("Height", 7),
                        ("Angle", 10), ("Distance", 9), ("Lighting", 9), ("Seq", 6), ("Type", 6)]:
            ctk.CTkLabel(hdr, text=txt, font=("Courier New", 9, "bold"),
                         text_color=ACCENT, width=w*7, anchor="w").pack(side="left", padx=4)

        self._preview_box = ctk.CTkTextbox(parent, font=("Courier New", 10),
                                           fg_color=BG_FIELD, text_color=TEXT_MAIN,
                                           border_color=BORDER, corner_radius=0)
        self._preview_box.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    # ── Shared helpers ────────────────────────────────────────────────────────
    def _section_label(self, parent, text, padx=0):
        ctk.CTkLabel(parent, text=text, font=("Courier New", 9, "bold"),
                     text_color=ACCENT).pack(anchor="w", padx=padx, pady=(8, 2))

    def _log(self, widget, text: str):
        widget.configure(state="normal")
        widget.insert("end", text + "\n")
        widget.see("end")
        widget.configure(state="disabled")

    def _clear_log(self, widget):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.configure(state="disabled")

    def _get_root(self) -> str | None:
        root = self.dataset_path.get().strip()
        if not root or not os.path.isdir(root):
            messagebox.showerror("Error", "Please select a valid dataset root folder.")
            return None
        return root

    # ══════════════════════════════════════════════════════════════════════════
    #  RENAME LOGIC
    # ══════════════════════════════════════════════════════════════════════════
    def _gather_rename_plan(self, dry=True) -> list[tuple[str, str]]:
        root = self._get_root()
        if not root:
            return []

        rv  = self._rename_vars
        rfv = self._rfilter_vars

        use_range = self._seq_mode.get() == "selected"
        seq_s = self._seq_start.get().strip() if use_range else ""
        seq_e = self._seq_end.get().strip()   if use_range else ""

        plan: list[tuple[str, str]] = []
        seen_new: set[str] = set()

        for abs_path, rel_folder, fname in walk_images(root):
            parts = parse_filename(fname)
            if not parts:
                continue

            skip = False
            for field, key in [("room","rf_room"),("height","rf_height"),
                                ("angle","rf_angle"),("distance","rf_dist"),
                                ("lighting","rf_light")]:
                v = rfv[key].get()
                if v and v != "(any)" and parts[field] != v:
                    skip = True; break
            if skip:
                continue

            if seq_s and int(parts["sequence"]) < int(seq_s): continue
            if seq_e and int(parts["sequence"]) > int(seq_e): continue

            new = dict(parts)
            for field, new_key in [("room","new_room"),("height","new_height"),
                                    ("angle","new_angle"),("distance","new_dist"),
                                    ("lighting","new_light")]:
                val = rv[new_key].get()
                if val and val != "(keep)":
                    new[field] = val

            new_name = build_filename(new)
            new_abs  = os.path.join(os.path.dirname(abs_path), new_name)
            if new_name == fname or new_abs in seen_new:
                continue
            seen_new.add(new_abs)
            plan.append((abs_path, new_abs))

        return plan

    def _preview_rename(self):
        plan = self._gather_rename_plan(dry=True)
        self._clear_log(self._rename_log)
        if not plan:
            self._log(self._rename_log, "No matching files found.")
            return
        self._log(self._rename_log, f"{'OLD':<55}  →  NEW")
        self._log(self._rename_log, "─" * 110)
        for old, new in plan:
            self._log(self._rename_log,
                      f"  {os.path.basename(old):<53}  →  {os.path.basename(new)}")
        self._log(self._rename_log, f"\n{len(plan)} file(s) would be renamed.")
        self._set_status(f"Preview: {len(plan)} rename operations")

    def _apply_rename(self):
        plan = self._gather_rename_plan(dry=False)
        if not plan:
            messagebox.showinfo("Nothing to do", "No matching files to rename.")
            return
        if not messagebox.askyesno("Confirm",
                                   f"Rename {len(plan)} file(s)?\n\n"
                                   "This cannot be undone unless you enabled backup."):
            return
        self._clear_log(self._rename_log)
        renamed = errors = 0
        for old, new in plan:
            try:
                if self._backup.get():
                    shutil.copy2(old, old + ".bak")
                os.rename(old, new)
                self._log(self._rename_log,
                          f"✔  {os.path.basename(old)}  →  {os.path.basename(new)}")
                renamed += 1
            except Exception as e:
                self._log(self._rename_log, f"✘  {os.path.basename(old)}  ERROR: {e}")
                errors += 1
        summary = f"Done: {renamed} renamed, {errors} error(s)."
        self._log(self._rename_log, "\n" + summary)
        self._set_status(summary)
        messagebox.showinfo("Complete", summary)

    # ══════════════════════════════════════════════════════════════════════════
    #  FILTER LOGIC
    # ══════════════════════════════════════════════════════════════════════════
    def _run_filter(self):
        root = self._get_root()
        if not root:
            return
        self._filter_count.set("Searching…")
        self._run_in_thread(self._run_filter_worker)

    def _run_filter_worker(self):
        root = self.dataset_path.get().strip()
        if not root or not os.path.isdir(root):
            return
        fv  = self._filter_vars
        ext = self._filter_ext.get()

        use_range = self._fseq_mode.get() == "selected"
        seq_s = self._fseq_start.get().strip() if use_range else ""
        seq_e = self._fseq_end.get().strip()   if use_range else ""

        # ── 1. Collect all matching files, group by base key ──────────────────
        # base_key = room_height_angle_distance_lighting_sequence  (no ext)
        grouped: dict[str, dict] = {}   # base_key -> {parts, color_path, depth_path}

        for abs_path, rel_folder, fname in walk_images(root):
            parts = parse_filename(fname)
            if not parts:
                continue

            skip = False
            for field, key in [("room","f_room"),("height","f_height"),
                                ("angle","f_angle"),("distance","f_dist"),
                                ("lighting","f_light")]:
                v = fv[key].get().strip()
                if v and v != "(any)" and parts[field] != v:
                    skip = True; break
            if skip:
                continue

            if seq_s and int(parts["sequence"]) < int(seq_s): continue
            if seq_e and int(parts["sequence"]) > int(seq_e): continue

            base_key = (f"{parts['room']}_{parts['height']}_{parts['angle']}_"
                        f"{parts['distance']}_{parts['lighting']}_{parts['sequence']}")

            if base_key not in grouped:
                grouped[base_key] = {"base_key": base_key,
                                     "parts": parts,
                                     "color_path": None,
                                     "depth_path": None}
            if parts["ext"] == ".jpg":
                grouped[base_key]["color_path"] = abs_path
            else:
                grouped[base_key]["depth_path"] = abs_path

        # ── 2. Apply extension filter, then sort by sequence ascending ──────────
        matches = []
        for rec in grouped.values():
            if ext == "jpg"  and not rec["color_path"]: continue
            if ext == "png"  and not rec["depth_path"]: continue
            matches.append(rec)

        matches.sort(key=lambda r: int(r["parts"]["sequence"]))

        # ── 3. Schedule UI rebuild on main thread ─────────────────────────────
        self.after(0, lambda m=matches: self._build_filter_rows(m))

    def _build_filter_rows(self, matches):
        """Must be called on the main thread. Rebuilds clickable result rows."""
        self._filter_matches = matches

        for w in self._filter_row_widgets:
            w.destroy()
        self._filter_row_widgets.clear()

        total = len(self._filter_matches)
        self._filter_count.set(f"{total} unique record(s)")

        for idx, rec in enumerate(self._filter_matches):
            p = rec["parts"]
            has = []
            if rec["color_path"]: has.append("JPG")
            if rec["depth_path"]: has.append("PNG")
            has_txt = "+".join(has)

            base_name = rec["base_key"]
            row_bg    = BG_FIELD if idx % 2 == 0 else BG_CARD

            row_frame = ctk.CTkFrame(self._filter_scroll,
                                     fg_color=row_bg, corner_radius=4,
                                     cursor="hand2")
            row_frame.pack(fill="x", pady=1)
            self._filter_row_widgets.append(row_frame)

            # columns
            cols = [
                (base_name,                                         44),
                (p["room"],                                          8),
                (p["height"],                                        7),
                (ANGLE_LABEL.get(p["angle"], p["angle"]),           10),
                (p["distance"],                                       9),
                (LIGHT_LABEL.get(p["lighting"], p["lighting"]),      8),
                (p["sequence"],                                       6),
                (has_txt,                                             7),
            ]
            for col_txt, col_w in cols:
                ctk.CTkLabel(row_frame, text=col_txt,
                             font=("Courier New", 10), text_color=TEXT_MAIN,
                             width=col_w*7, anchor="w").pack(side="left", padx=4, pady=4)

            # bind click on the whole row and all its children
            rec_copy = dict(rec)
            def _on_click(event, i=idx):
                self._open_image_picker(i)
            row_frame.bind("<Button-1>", _on_click)
            for child in row_frame.winfo_children():
                child.bind("<Button-1>", _on_click)

            # hover highlight
            def _enter(e, f=row_frame, bg=row_bg):
                f.configure(fg_color=BG_MID)
                for c in f.winfo_children(): c.configure(fg_color=BG_MID)
            def _leave(e, f=row_frame, bg=row_bg):
                f.configure(fg_color=bg)
                for c in f.winfo_children(): c.configure(fg_color=bg)
            row_frame.bind("<Enter>", _enter)
            row_frame.bind("<Leave>", _leave)
            for child in row_frame.winfo_children():
                child.bind("<Enter>", _enter)
                child.bind("<Leave>", _leave)

        self._set_status(f"Filter: {total} unique record(s)")

    def _open_image_picker(self, start_idx: int):
        """Full inline image viewer with object counter panel and Excel export."""
        try:
            from PIL import Image, ImageTk
            _pil_ok = True
        except ImportError:
            _pil_ok = False

        matches = self._filter_matches
        total   = len(matches)
        if total == 0:
            return

        IMG_W, IMG_H = 900, 560

        state = {
            "idx":   start_idx,
            "mode":  "color",
            "imgtk": None,
            # object_counts: { image_base_key -> { obj_name -> count } }
            "object_counts": {},
        }

        win = ctk.CTkToplevel(self)
        win.title("Image Viewer")
        win.geometry(f"{IMG_W + 300}x{IMG_H + 160}")
        win.minsize(900, 600)
        win.resizable(True, True)
        win.configure(fg_color=BG_DARK)
        win.grab_set()
        win.lift()
        win.focus_force()

        # ── TOP BAR ──────────────────────────────────────────────────────────
        top = ctk.CTkFrame(win, fg_color=BG_MID, corner_radius=0, height=38)
        top.pack(fill="x")
        top.pack_propagate(False)
        counter_lbl = ctk.CTkLabel(top, text="", font=("Courier New", 10),
                                   text_color=TEXT_DIM)
        counter_lbl.pack(side="left", padx=14)
        key_lbl = ctk.CTkLabel(top, text="", font=("Courier New", 10, "bold"),
                               text_color=TEXT_MAIN)
        key_lbl.pack(side="left", padx=8)
        detail_lbl = ctk.CTkLabel(top, text="", font=("Courier New", 9),
                                  text_color=TEXT_DIM)
        detail_lbl.pack(side="right", padx=16)
        seq_lbl = ctk.CTkLabel(top, text="", font=("Courier New", 10),
                               text_color=ACCENT)
        seq_lbl.pack(side="right", padx=14)

        # ── MAIN BODY: image left, panel right ────────────────────────────────
        body = ctk.CTkFrame(win, fg_color=BG_DARK, corner_radius=0)
        body.pack(fill="both", expand=True)

        # ── LEFT: image canvas ────────────────────────────────────────────────
        canvas_frame = ctk.CTkFrame(body, fg_color=BG_FIELD, corner_radius=0)
        canvas_frame.pack(side="left", fill="both", expand=True)

        canvas = tk.Canvas(canvas_frame, bg="#0F1117", highlightthickness=0)
        canvas.pack(fill="both", expand=True)

        no_img_lbl = ctk.CTkLabel(canvas_frame, text="No image available",
                                  font=("Courier New", 13), text_color=TEXT_DIM,
                                  fg_color="transparent")

        # ── RIGHT: object counter panel ───────────────────────────────────────
        right_panel = ctk.CTkFrame(body, fg_color=BG_CARD, corner_radius=0, width=280)
        right_panel.pack(side="right", fill="y")
        right_panel.pack_propagate(False)

        ctk.CTkLabel(right_panel, text="OBJECT  COUNTER",
                     font=("Courier New", 9, "bold"), text_color=ACCENT).pack(
                         anchor="w", padx=14, pady=(12, 4))

        add_obj_btn = ctk.CTkButton(right_panel, text="+ Add Object", height=32,
                                    fg_color=ACCENT2, hover_color=ACCENT,
                                    font=("Courier New", 11, "bold"), text_color="white",
                                    command=lambda: _add_object_dialog())
        add_obj_btn.pack(fill="x", padx=12, pady=(0, 8))

        # scrollable table area
        obj_scroll = ctk.CTkScrollableFrame(right_panel, fg_color=BG_FIELD,
                                            corner_radius=6)
        obj_scroll.pack(fill="both", expand=True, padx=12, pady=(0, 6))

        # list of (obj_name, count_var, row_frame) — rebuilt per image
        obj_rows: list[dict] = []

        def _get_counts_for_current():
            key = matches[state["idx"]]["base_key"]
            if key not in state["object_counts"]:
                state["object_counts"][key] = {}
            return state["object_counts"][key]

        def _rebuild_obj_table():
            for w in obj_scroll.winfo_children():
                w.destroy()
            obj_rows.clear()

            counts = _get_counts_for_current()
            # ensure all known objects appear even if count=0
            all_objs = _all_known_objects()
            for name in all_objs:
                if name not in counts:
                    counts[name] = 0
                _add_obj_row(name, counts)

            if not all_objs:
                ctk.CTkLabel(obj_scroll,
                             text="No objects added yet.\nClick '+ Add Object'.",
                             font=("Courier New", 9), text_color=TEXT_DIM,
                             justify="center").pack(pady=20)

        def _all_known_objects():
            """Return the union of all object names ever added."""
            names = set()
            for d in state["object_counts"].values():
                names.update(d.keys())
            return sorted(names)

        def _add_obj_row(name: str, counts: dict):
            row = ctk.CTkFrame(obj_scroll, fg_color=BG_MID, corner_radius=4)
            row.pack(fill="x", pady=2)

            ctk.CTkLabel(row, text=name, font=("Courier New", 10),
                         text_color=TEXT_MAIN, anchor="w", width=100).pack(
                             side="left", padx=8, pady=6)

            cnt_var = tk.StringVar(value=str(counts.get(name, 0)))

            minus_btn = ctk.CTkButton(row, text="−", width=28, height=28,
                                      fg_color=BG_FIELD, hover_color=DANGER,
                                      font=("Courier New", 13, "bold"),
                                      text_color=TEXT_MAIN, corner_radius=4,
                                      command=lambda n=name: _change_count(n, -1))
            minus_btn.pack(side="left", padx=2)

            ctk.CTkLabel(row, textvariable=cnt_var,
                         font=("Courier New", 12, "bold"),
                         text_color=ACCENT, width=34,
                         anchor="center").pack(side="left", padx=2)

            plus_btn = ctk.CTkButton(row, text="+", width=28, height=28,
                                     fg_color=BG_FIELD, hover_color=SUCCESS,
                                     font=("Courier New", 13, "bold"),
                                     text_color=TEXT_MAIN, corner_radius=4,
                                     command=lambda n=name: _change_count(n, +1))
            plus_btn.pack(side="left", padx=2)

            obj_rows.append({"name": name, "var": cnt_var})

        def _change_count(name: str, delta: int):
            counts = _get_counts_for_current()
            counts[name] = max(0, counts.get(name, 0) + delta)
            # refresh just the var
            for r in obj_rows:
                if r["name"] == name:
                    r["var"].set(str(counts[name]))
                    break

        def _add_object_dialog():
            dlg = ctk.CTkToplevel(win)
            dlg.title("Add Object")
            dlg.geometry("320x160")
            dlg.resizable(False, False)
            dlg.configure(fg_color=BG_DARK)
            dlg.grab_set()
            dlg.lift()
            dlg.focus_force()

            ctk.CTkLabel(dlg, text="Object name:", font=("Courier New", 11),
                         text_color=TEXT_MAIN).pack(pady=(20, 6))
            name_entry = ctk.CTkEntry(dlg, width=220, font=("Courier New", 12),
                                      fg_color=BG_FIELD, text_color=TEXT_MAIN,
                                      border_color=ACCENT,
                                      placeholder_text="e.g. chair")
            name_entry.pack(pady=4)
            name_entry.focus_set()

            def _save():
                name = name_entry.get().strip()
                if not name:
                    return
                # add to every existing image record so table is consistent
                for d in state["object_counts"].values():
                    if name not in d:
                        d[name] = 0
                # also seed current image
                counts = _get_counts_for_current()
                if name not in counts:
                    counts[name] = 0
                dlg.destroy()
                _rebuild_obj_table()

            ctk.CTkButton(dlg, text="Save", fg_color=ACCENT, hover_color=ACCENT2,
                          font=("Courier New", 11, "bold"),
                          command=_save).pack(pady=12)
            dlg.bind("<Return>", lambda e: _save())

        # ── BOTTOM BAR ────────────────────────────────────────────────────────
        bot = ctk.CTkFrame(win, fg_color=BG_MID, corner_radius=0, height=96)
        bot.pack(fill="x", side="bottom")
        bot.pack_propagate(False)

        nav_row = ctk.CTkFrame(bot, fg_color="transparent")
        nav_row.pack(pady=(10, 4))

        prev_btn = ctk.CTkButton(nav_row, text="◀", width=52, height=36,
                                 fg_color=BG_FIELD, border_width=1, border_color=BORDER,
                                 hover_color=BG_CARD, font=("Courier New", 14, "bold"),
                                 text_color=TEXT_MAIN)
        prev_btn.pack(side="left", padx=6)

        color_tab = ctk.CTkButton(nav_row, text="🖼  Color", width=120, height=36,
                                  fg_color=ACCENT, hover_color=ACCENT2,
                                  font=("Courier New", 11, "bold"), text_color="white")
        color_tab.pack(side="left", padx=4)

        depth_tab = ctk.CTkButton(nav_row, text="◧  Depth", width=120, height=36,
                                  fg_color=BG_FIELD, border_width=1, border_color=ACCENT2,
                                  hover_color=BG_CARD, font=("Courier New", 11),
                                  text_color=ACCENT2)
        depth_tab.pack(side="left", padx=4)

        next_btn = ctk.CTkButton(nav_row, text="▶", width=52, height=36,
                                 fg_color=BG_FIELD, border_width=1, border_color=BORDER,
                                 hover_color=BG_CARD, font=("Courier New", 14, "bold"),
                                 text_color=TEXT_MAIN)
        next_btn.pack(side="left", padx=6)

        save_btn = ctk.CTkButton(nav_row, text="💾  Save Record", width=150, height=36,
                                 fg_color=SUCCESS, hover_color="#2AB87A",
                                 font=("Courier New", 11, "bold"), text_color="white",
                                 command=lambda: _save_excel())
        save_btn.pack(side="left", padx=14)

        fname_lbl = ctk.CTkLabel(bot, text="", font=("Courier New", 9),
                                 text_color=TEXT_DIM)
        fname_lbl.pack(pady=(0, 6))

        # ── IMAGE RENDER ──────────────────────────────────────────────────────
        def _render_image(path):
            no_img_lbl.place_forget()
            canvas.delete("all")
            if not path or not os.path.isfile(path):
                no_img_lbl.place(relx=0.5, rely=0.5, anchor="center")
                return
            if not _pil_ok:
                no_img_lbl.configure(
                    text="Install Pillow to preview:\npip install Pillow")
                no_img_lbl.place(relx=0.5, rely=0.5, anchor="center")
                return
            try:
                img = Image.open(path)
                cw = canvas.winfo_width()  or IMG_W
                ch = canvas.winfo_height() or IMG_H
                img.thumbnail((cw, ch), Image.LANCZOS)
                imgtk = ImageTk.PhotoImage(img)
                state["imgtk"] = imgtk
                canvas.create_image(cw // 2, ch // 2, anchor="center", image=imgtk)
            except Exception as ex:
                no_img_lbl.configure(text=f"Cannot load image:\n{ex}")
                no_img_lbl.place(relx=0.5, rely=0.5, anchor="center")

        def _highlight_tabs():
            if state["mode"] == "color":
                color_tab.configure(fg_color=ACCENT, border_width=0, text_color="white")
                depth_tab.configure(fg_color=BG_FIELD, border_width=1,
                                    border_color=ACCENT2, text_color=ACCENT2)
            else:
                color_tab.configure(fg_color=BG_FIELD, border_width=1,
                                    border_color=ACCENT, text_color=ACCENT)
                depth_tab.configure(fg_color=ACCENT2, border_width=0, text_color="white")

        # ── REFRESH ───────────────────────────────────────────────────────────
        def refresh():
            i   = state["idx"]
            rec = matches[i]
            p   = rec["parts"]

            has_color = bool(rec["color_path"] and os.path.isfile(rec["color_path"]))
            has_depth = bool(rec["depth_path"] and os.path.isfile(rec["depth_path"]))

            if state["mode"] == "color" and not has_color and has_depth:
                state["mode"] = "depth"
            elif state["mode"] == "depth" and not has_depth and has_color:
                state["mode"] = "color"

            counter_lbl.configure(text=f"{i+1} / {total}")
            key_lbl.configure(text=rec["base_key"])
            seq_lbl.configure(text=f"seq {p['sequence']}")
            detail_lbl.configure(
                text=(f"Room {p['room']}  ·  {p['height']}  ·  "
                      f"{ANGLE_LABEL.get(p['angle'], p['angle'])}  ·  "
                      f"{p['distance']}  ·  "
                      f"{LIGHT_LABEL.get(p['lighting'], p['lighting'])}"))

            color_tab.configure(state="normal" if has_color else "disabled")
            depth_tab.configure(state="normal" if has_depth else "disabled")
            _highlight_tabs()

            path = rec["color_path"] if state["mode"] == "color" else rec["depth_path"]
            _render_image(path)
            fname_lbl.configure(text=os.path.basename(path) if path else "—")

            prev_btn.configure(state="normal" if i > 0 else "disabled",
                               text_color=TEXT_MAIN if i > 0 else TEXT_DIM)
            next_btn.configure(state="normal" if i < total-1 else "disabled",
                               text_color=TEXT_MAIN if i < total-1 else TEXT_DIM)

            _rebuild_obj_table()

        # ── SAVE TO EXCEL ─────────────────────────────────────────────────────
        def _save_excel():
            if not _XLSX_OK:
                messagebox.showerror(
                    "Missing library",
                    "openpyxl is required to save Excel files.\n"
                    "Run: pip install openpyxl")
                return

            all_objs = _all_known_objects()
            if not state["object_counts"]:
                messagebox.showinfo("Nothing to save",
                                    "No images visited yet.")
                return

            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel file", "*.xlsx")],
                title="Save object record…",
                parent=win)
            if not path:
                return

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Object Counts"

            # ── colour palette ────────────────────────────────────────────────
            C_HDR_BG   = "1E2535"
            C_HDR_FG   = "4F8EF7"
            C_OBJ_BG   = "252D3D"
            C_OBJ_FG   = "7C5CFC"
            C_ROW_EVEN = "161B27"
            C_ROW_ODD  = "1E2535"
            C_BORDER   = "2E3A50"

            hdr_font  = Font(name="Calibri", bold=True,  color=C_HDR_FG, size=10)
            obj_font  = Font(name="Calibri", bold=True,  color=C_OBJ_FG, size=10)
            data_font = Font(name="Calibri", bold=False, color="E8EDF5",  size=10)
            cnt_font  = Font(name="Calibri", bold=True,  color="E8EDF5",  size=10)

            hdr_fill  = PatternFill("solid", fgColor=C_HDR_BG)
            obj_fill  = PatternFill("solid", fgColor=C_OBJ_BG)

            center = Alignment(horizontal="center", vertical="center",
                               wrap_text=True)
            left   = Alignment(horizontal="left",   vertical="center",
                               wrap_text=False)

            thin   = Side(style="thin",   color=C_BORDER)
            medium = Side(style="medium", color="4F8EF7")
            brd    = Border(left=thin, right=thin, top=thin, bottom=thin)
            brd_m  = Border(left=medium, right=thin, top=thin, bottom=thin)

            # ── column definitions ────────────────────────────────────────────
            # (header_text, width, alignment, is_meta)
            META_COLS = [
                ("Date",                    13, left,   True),
                ("Floor",                    7, center, True),
                ("Room",                     7, center, True),
                ("Height (m)",               9, center, True),
                ("Distance",                10, center, True),
                ("Angle",                   11, center, True),
                ("Lighting",                10, center, True),
                ("Resolution",              12, center, True),
                ("RGB Format",              10, center, True),
                ("Depth Format",            11, center, True),
                ("Start Filename",          34, left,   True),
                ("End Filename",            34, left,   True),
                ("# Images",                9, center, True),
                ("Est. Total Objects",      14, center, True),
            ]
            # object columns (dynamic)
            OBJ_COLS = [(name, max(12, len(name)+2), center, False)
                        for name in all_objs]
            # trailing meta
            TRAIL_COLS = [
                ("Object Class",            18, left,   True),
                ("Notes",                   28, left,   True),
            ]

            ALL_COLS = META_COLS + OBJ_COLS + TRAIL_COLS
            N_META   = len(META_COLS)
            N_OBJ    = len(OBJ_COLS)

            # ── header row ────────────────────────────────────────────────────
            for ci, (hdr, width, align, is_meta) in enumerate(ALL_COLS, start=1):
                cell = ws.cell(row=1, column=ci, value=hdr)
                if not is_meta:                       # object column
                    cell.font      = obj_font
                    cell.fill      = obj_fill
                else:
                    cell.font      = hdr_font
                    cell.fill      = hdr_fill
                cell.alignment = center               # all headers centred
                cell.border    = brd
                ws.column_dimensions[
                    openpyxl.utils.get_column_letter(ci)].width = width

            ws.row_dimensions[1].height = 30

            # ── data rows ─────────────────────────────────────────────────────
            import datetime

            sorted_keys = sorted(
                state["object_counts"].keys(),
                key=lambda k: int(
                    next((m["parts"]["sequence"]
                          for m in matches if m["base_key"] == k), "0")))

            for row_i, key in enumerate(sorted_keys, start=2):
                rec    = next((m for m in matches if m["base_key"] == key), None)
                counts = state["object_counts"].get(key, {})

                row_fill = PatternFill("solid",
                                       fgColor=C_ROW_EVEN if row_i % 2 == 0
                                       else C_ROW_ODD)

                if rec:
                    p = rec["parts"]
                    floor_code = p["room"][:2]
                    room_code  = p["room"][2:]
                    # derive start/end filenames from color files if available
                    start_f = (os.path.basename(rec["color_path"])
                               if rec["color_path"] else "")
                    end_f   = (os.path.basename(rec["color_path"])
                               if rec["color_path"] else "")

                    n_images = 1   # each row = one image pair

                    total_objs = sum(counts.values())

                    # object class = comma-joined names that have count > 0
                    obj_class = ", ".join(
                        n for n in all_objs if counts.get(n, 0) > 0)

                    meta_vals = [
                        datetime.date.today().isoformat(),  # Date
                        floor_code,                          # Floor
                        room_code,                           # Room
                        p["height"].replace("m", ""),        # Height
                        p["distance"].capitalize(),          # Distance
                        ANGLE_LABEL.get(p["angle"], p["angle"]),  # Angle
                        LIGHT_LABEL.get(p["lighting"], p["lighting"]),  # Lighting
                        "1280x720",                          # Resolution (default)
                        "jpg",                               # RGB Format
                        "png",                               # Depth Format
                        start_f,                             # Start Filename
                        end_f,                               # End Filename
                        n_images,                            # # Images
                        total_objs,                          # Est. Total Objects
                    ]
                else:
                    meta_vals = [
                        datetime.date.today().isoformat(),
                        "", "", "", "", "", "", "", "", "",
                        key, "", 1, 0]
                    obj_class = ""

                obj_vals   = [counts.get(o, 0) for o in all_objs]
                trail_vals = [obj_class, ""]

                all_vals = meta_vals + obj_vals + trail_vals

                for ci, (val, (_, _, align, is_meta)) in enumerate(
                        zip(all_vals, ALL_COLS), start=1):
                    cell = ws.cell(row=row_i, column=ci, value=val)
                    cell.fill   = row_fill
                    cell.border = brd
                    if not is_meta:
                        cell.font      = cnt_font
                        cell.alignment = center
                    else:
                        cell.font      = data_font
                        cell.alignment = align

                ws.row_dimensions[row_i].height = 18

            ws.freeze_panes = "A2"

            # ── totals row ────────────────────────────────────────────────────
            tot_row = len(sorted_keys) + 2
            tot_fill = PatternFill("solid", fgColor="0F1117")
            tot_font = Font(name="Calibri", bold=True, color="4F8EF7", size=10)

            for ci in range(1, len(ALL_COLS)+1):
                cell = ws.cell(row=tot_row, column=ci)
                cell.fill   = tot_fill
                cell.border = brd
                cell.font   = tot_font
                col_name = ALL_COLS[ci-1][0]
                if col_name == "Date":
                    cell.value     = "TOTAL"
                    cell.alignment = left
                elif col_name == "# Images":
                    cell.value     = len(sorted_keys)
                    cell.alignment = center
                elif col_name == "Est. Total Objects":
                    cell.value     = sum(
                        sum(state["object_counts"].get(k, {}).values())
                        for k in sorted_keys)
                    cell.alignment = center
                elif N_META <= ci-1 < N_META + N_OBJ:
                    # sum each object column
                    obj_name = all_objs[ci-1-N_META]
                    cell.value = sum(
                        state["object_counts"].get(k, {}).get(obj_name, 0)
                        for k in sorted_keys)
                    cell.alignment = center

            ws.row_dimensions[tot_row].height = 20

            try:
                wb.save(path)
                messagebox.showinfo(
                    "Saved",
                    f"Record saved to:\n{path}\n\n"
                    f"{len(sorted_keys)} row(s)  ·  "
                    f"{len(all_objs)} object column(s)",
                    parent=win)
            except Exception as e:
                messagebox.showerror("Save failed", str(e), parent=win)

        # ── wire up buttons ───────────────────────────────────────────────────
        def set_mode(m):
            state["mode"] = m
            refresh()

        def go_prev():
            if state["idx"] > 0:
                state["idx"] -= 1
                refresh()

        def go_next():
            if state["idx"] < total - 1:
                state["idx"] += 1
                refresh()

        color_tab.configure(command=lambda: set_mode("color"))
        depth_tab.configure(command=lambda: set_mode("depth"))
        prev_btn.configure(command=go_prev)
        next_btn.configure(command=go_next)

        win.bind("<Left>",  lambda e: go_prev())
        win.bind("<Right>", lambda e: go_next())
        win.bind("c",       lambda e: set_mode("color"))
        win.bind("d",       lambda e: set_mode("depth"))

        canvas.bind("<Configure>",
                    lambda e: win.after(80, refresh) if e.widget == canvas else None)

        win.after(100, refresh)

    def _export_filter_list(self):
        if not self._filter_matches:
            messagebox.showinfo("Empty", "Run a search first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text file", "*.txt"), ("All", "*.*")])
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write("base_key\troom\theight\tangle\tdistance\tlighting\tsequence"
                    "\tcolor_path\tdepth_path\n")
            for rec in self._filter_matches:
                p = rec["parts"]
                f.write(f"{rec['base_key']}\t{p['room']}\t{p['height']}\t"
                        f"{p['angle']}\t{p['distance']}\t{p['lighting']}\t"
                        f"{p['sequence']}\t{rec['color_path'] or ''}\t"
                        f"{rec['depth_path'] or ''}\n")
        messagebox.showinfo("Exported",
                            f"Saved {len(self._filter_matches)} rows to\n{path}")

    def _copy_matched(self):
        if not self._filter_matches:
            messagebox.showinfo("Empty", "Run a search first.")
            return
        dest = filedialog.askdirectory(title="Copy matched files to…")
        if not dest:
            return
        copied = 0
        for rec in self._filter_matches:
            for path in [rec["color_path"], rec["depth_path"]]:
                if path and os.path.isfile(path):
                    try:
                        shutil.copy2(path, dest)
                        copied += 1
                    except Exception:
                        pass
        messagebox.showinfo("Done", f"Copied {copied} file(s) to\n{dest}")

    # ══════════════════════════════════════════════════════════════════════════
    #  THREADING HELPERS
    # ══════════════════════════════════════════════════════════════════════════
    def _run_in_thread(self, fn, *args):
        t = threading.Thread(target=fn, args=args, daemon=True)
        t.start()

    def _safe_ui(self, fn, *args):
        self.after(0, fn, *args)

    # ══════════════════════════════════════════════════════════════════════════
    #  PREVIEW / SCAN  (threaded)
    # ══════════════════════════════════════════════════════════════════════════
    def _scan_all(self):
        root = self._get_root()
        if not root:
            return
        self._clear_log(self._preview_box)
        self._scan_count.set("Scanning…")
        self._set_status("Scanning…")

        def _worker():
            rows = []
            total = bad = 0
            for abs_path, rel_folder, fname in walk_images(root):
                parts = parse_filename(fname)
                if parts:
                    ftype = "color" if parts["ext"] == ".jpg" else "depth"
                    rows.append(
                        f"  {rel_folder:<30}  {parts['room']:8}  {parts['height']:7}  "
                        f"{ANGLE_LABEL.get(parts['angle'],parts['angle']):10}  "
                        f"{parts['distance']:9}  "
                        f"{LIGHT_LABEL.get(parts['lighting'],parts['lighting']):8}  "
                        f"{parts['sequence']:6}  {ftype}")
                    total += 1
                else:
                    rows.append(f"  [UNRECOGNISED]  {rel_folder}/{fname}")
                    bad += 1

            def _update():
                self._preview_box.configure(state="normal")
                for r in rows:
                    self._preview_box.insert("end", r + "\n")
                self._preview_box.configure(state="disabled")
                self._scan_count.set(f"{total} valid, {bad} unrecognised")
                self._set_status(f"Scan complete: {total} images found")
            self.after(0, _update)

        self._run_in_thread(_worker)


# ─── Entry ────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = DatasetManagerApp()
    app.mainloop()
