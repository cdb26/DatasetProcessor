import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
import shutil
from collections import defaultdict

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

        # ── 2. Apply extension filter to decide which records to show ─────────
        self._filter_matches = []
        for rec in grouped.values():
            if ext == "jpg"  and not rec["color_path"]: continue
            if ext == "png"  and not rec["depth_path"]: continue
            self._filter_matches.append(rec)

        # ── 3. Clear old row widgets and rebuild ──────────────────────────────
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
            def _on_click(event, r=rec_copy):
                self._open_image_picker(r)
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

    def _open_image_picker(self, rec: dict):
        """Popup that lets the user open the color (.jpg), depth (.png), or both."""
        has_color = rec["color_path"] and os.path.isfile(rec["color_path"])
        has_depth = rec["depth_path"] and os.path.isfile(rec["depth_path"])

        if not has_color and not has_depth:
            messagebox.showerror("Not found", "Neither image file could be located on disk.")
            return

        win = ctk.CTkToplevel(self)
        win.title("Open image")
        win.geometry("380x260")
        win.resizable(False, False)
        win.configure(fg_color=BG_DARK)
        win.grab_set()
        win.lift()
        win.focus_force()

        p = rec["parts"]
        title_txt = (f"{rec['base_key']}")
        ctk.CTkLabel(win, text="Open image for:", font=("Courier New", 10),
                     text_color=TEXT_DIM).pack(pady=(18, 0))
        ctk.CTkLabel(win, text=title_txt, font=("Courier New", 12, "bold"),
                     text_color=TEXT_MAIN).pack(pady=(2, 14))

        detail = (f"Room {p['room']}  ·  {p['height']}  ·  "
                  f"{ANGLE_LABEL.get(p['angle'], p['angle'])}  ·  "
                  f"{p['distance']}  ·  {LIGHT_LABEL.get(p['lighting'], p['lighting'])}  ·  "
                  f"seq {p['sequence']}")
        ctk.CTkLabel(win, text=detail, font=("Courier New", 9),
                     text_color=TEXT_DIM).pack(pady=(0, 16))

        btn_frame = ctk.CTkFrame(win, fg_color="transparent")
        btn_frame.pack(pady=4)

        def _open(path):
            try:
                os.startfile(path)
            except AttributeError:
                import subprocess, sys
                opener = "open" if sys.platform == "darwin" else "xdg-open"
                subprocess.Popen([opener, path])
            except Exception as e:
                messagebox.showerror("Cannot open", str(e))
            win.destroy()

        def _open_both():
            if has_color: _open_file(rec["color_path"])
            if has_depth: _open_file(rec["depth_path"])
            win.destroy()

        def _open_file(path):
            try:
                os.startfile(path)
            except AttributeError:
                import subprocess, sys
                opener = "open" if sys.platform == "darwin" else "xdg-open"
                subprocess.Popen([opener, path])
            except Exception as e:
                messagebox.showerror("Cannot open", str(e))

        if has_color:
            ctk.CTkButton(btn_frame, text="🖼  Color  (.jpg)", width=150, height=42,
                          fg_color=ACCENT, hover_color=ACCENT2,
                          font=("Courier New", 12, "bold"), text_color="white",
                          command=lambda: (_open_file(rec["color_path"]), win.destroy())
                          ).pack(side="left", padx=8)
        if has_depth:
            ctk.CTkButton(btn_frame, text="◧  Depth  (.png)", width=150, height=42,
                          fg_color=ACCENT2, hover_color=ACCENT,
                          font=("Courier New", 12, "bold"), text_color="white",
                          command=lambda: (_open_file(rec["depth_path"]), win.destroy())
                          ).pack(side="left", padx=8)
        if has_color and has_depth:
            ctk.CTkButton(win, text="Open both", width=120, height=32,
                          fg_color=BG_FIELD, border_width=1, border_color=BORDER,
                          hover_color=BG_MID, font=("Courier New", 10), text_color=TEXT_DIM,
                          command=_open_both).pack(pady=(10, 0))

        # Show file paths
        if has_color:
            ctk.CTkLabel(win, text=f"Color: …{os.sep}{os.path.basename(rec['color_path'])}",
                         font=("Courier New", 8), text_color=TEXT_DIM).pack(pady=(8, 0))
        if has_depth:
            ctk.CTkLabel(win, text=f"Depth: …{os.sep}{os.path.basename(rec['depth_path'])}",
                         font=("Courier New", 8), text_color=TEXT_DIM).pack(pady=(2, 0))

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
    #  PREVIEW / SCAN
    # ══════════════════════════════════════════════════════════════════════════
    def _scan_all(self):
        root = self._get_root()
        if not root:
            return
        self._clear_log(self._preview_box)
        total = bad = 0
        for abs_path, rel_folder, fname in walk_images(root):
            parts = parse_filename(fname)
            if parts:
                ftype = "color" if parts["ext"] == ".jpg" else "depth"
                row = (f"  {rel_folder:<30}  {parts['room']:8}  {parts['height']:7}  "
                       f"{ANGLE_LABEL.get(parts['angle'],parts['angle']):10}  "
                       f"{parts['distance']:9}  "
                       f"{LIGHT_LABEL.get(parts['lighting'],parts['lighting']):8}  "
                       f"{parts['sequence']:6}  {ftype}")
                self._log(self._preview_box, row)
                total += 1
            else:
                self._log(self._preview_box, f"  [UNRECOGNISED]  {rel_folder}/{fname}")
                bad += 1
        self._scan_count.set(f"{total} valid, {bad} unrecognised")
        self._set_status(f"Scan complete: {total} images found")


# ─── Entry ────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = DatasetManagerApp()
    app.mainloop()
