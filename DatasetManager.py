import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
import shutil
from pathlib import Path
from collections import defaultdict
import threading

# ─── Theme ───────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

ACCENT      = "#4F8EF7"
ACCENT2     = "#7C5CFC"
BG_DARK     = "#0F1117"
BG_MID      = "#161B27"
BG_CARD     = "#1E2535"
BG_FIELD    = "#252D3D"
TEXT_MAIN   = "#E8EDF5"
TEXT_DIM    = "#7A8599"
SUCCESS     = "#3DD68C"
WARNING     = "#F5A623"
DANGER      = "#F75C5C"
BORDER      = "#2E3A50"

# ─── Filename parser ──────────────────────────────────────────────────────────
FILENAME_RE = re.compile(
    r"^(\d{6})_([\d.]+m)_(\d)_(close|medium|far)_(dim|well)_(\d{4})(\.(?:jpg|png))$",
    re.IGNORECASE
)

def parse_filename(name: str) -> dict | None:
    m = FILENAME_RE.match(name)
    if not m:
        return None
    return {
        "room":      m.group(1),
        "height":    m.group(2),
        "angle":     m.group(3),
        "distance":  m.group(4),
        "lighting":  m.group(5),
        "sequence":  m.group(6),
        "ext":       m.group(7).lower(),
        "original":  name,
    }

def build_filename(parts: dict) -> str:
    return (f"{parts['room']}_{parts['height']}_{parts['angle']}_"
            f"{parts['distance']}_{parts['lighting']}_{parts['sequence']}{parts['ext']}")

ANGLE_LABEL   = {"1": "Ortho", "2": "Diagonal", "3": "Top-down"}
DIST_LABEL    = {"close": "Close", "medium": "Medium", "far": "Far"}
LIGHT_LABEL   = {"dim": "Dim", "well": "Well-lit"}
HEIGHT_OPTS   = ["0.8m", "1.2m", "1.6m"]
ANGLE_OPTS    = ["1", "2", "3"]
DIST_OPTS     = ["close", "medium", "far"]
LIGHT_OPTS    = ["dim", "well"]


# ─── Helpers ──────────────────────────────────────────────────────────────────
def walk_images(root: str):
    """Yield (abs_path, relative_folder, filename) for every image under root."""
    for dirpath, _, files in os.walk(root):
        for f in sorted(files):
            if f.lower().endswith((".jpg", ".png")):
                yield os.path.join(dirpath, f), os.path.relpath(dirpath, root), f


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN APP
# ═══════════════════════════════════════════════════════════════════════════════
class DatasetManagerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Dataset Manager")
        self.geometry("1200x780")
        self.minsize(960, 660)
        self.configure(fg_color=BG_DARK)

        self.dataset_path = tk.StringVar(value="")
        self.status_var   = tk.StringVar(value="Ready")
        self._build_ui()

    # ── Layout ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header
        hdr = ctk.CTkFrame(self, fg_color=BG_MID, corner_radius=0, height=60)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="⬡  DATASET MANAGER",
                     font=("Courier New", 17, "bold"),
                     text_color=ACCENT).pack(side="left", padx=24, pady=0)
        ctk.CTkLabel(hdr, text="Image filename toolkit for structured data collection",
                     font=("Courier New", 11),
                     text_color=TEXT_DIM).pack(side="left", padx=4)

        # ── Path row
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

        # ── Tab view
        self.tabs = ctk.CTkTabview(self, fg_color=BG_CARD,
                                   segmented_button_fg_color=BG_MID,
                                   segmented_button_selected_color=ACCENT,
                                   segmented_button_unselected_color=BG_MID,
                                   segmented_button_selected_hover_color=ACCENT2,
                                   text_color=TEXT_MAIN,
                                   corner_radius=8)
        self.tabs.pack(fill="both", expand=True, padx=16, pady=(10, 0))
        self.tabs.add("✦  Rename")
        self.tabs.add("⊞  Filter / Search")
        self.tabs.add("◈  Preview")

        self._build_rename_tab(self.tabs.tab("✦  Rename"))
        self._build_filter_tab(self.tabs.tab("⊞  Filter / Search"))
        self._build_preview_tab(self.tabs.tab("◈  Preview"))

        # ── Status bar
        sb = ctk.CTkFrame(self, fg_color=BG_MID, corner_radius=0, height=30)
        sb.pack(fill="x", side="bottom")
        sb.pack_propagate(False)
        ctk.CTkLabel(sb, textvariable=self.status_var,
                     font=("Courier New", 10), text_color=TEXT_DIM).pack(
                         side="left", padx=16)

    # ── Browse ────────────────────────────────────────────────────────────────
    def _browse(self):
        d = filedialog.askdirectory(title="Select dataset root folder")
        if d:
            self.dataset_path.set(d)
            self._set_status(f"Dataset root: {d}")

    def _set_status(self, msg: str, color: str = TEXT_DIM):
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

        # ── Left: field editors ───────────────────────────────────────────────
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
                mb = ctk.CTkOptionMenu(grid, variable=var, values=["(keep)"] + opts,
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

        # sequence range
        seq_row = ctk.CTkFrame(left, fg_color="transparent")
        seq_row.pack(fill="x", pady=(10, 0))
        self._section_label(seq_row, "SEQUENCE RANGE  (4 digits, e.g. 0617 – 0703)")
        seq_inner = ctk.CTkFrame(seq_row, fg_color=BG_FIELD, corner_radius=8)
        seq_inner.pack(fill="x")
        self._seq_start = tk.StringVar()
        self._seq_end   = tk.StringVar()
        for label, var in [("From", self._seq_start), ("To  ", self._seq_end)]:
            ctk.CTkLabel(seq_inner, text=label, font=("Courier New", 10),
                         text_color=TEXT_DIM).pack(side="left", padx=(14, 4), pady=8)
            ctk.CTkEntry(seq_inner, textvariable=var, width=90,
                         font=("Courier New", 11), fg_color=BG_CARD,
                         text_color=TEXT_MAIN, border_color=BORDER,
                         placeholder_text="0617").pack(side="left", padx=4, pady=8)

        # filter rows to rename
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
                         text_color=TEXT_DIM).grid(row=j//3, column=(j%3)*2,
                                                    sticky="w", padx=10, pady=6)
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
                                 row=j//3, column=(j%3)*2+1, sticky="w",
                                 padx=6, pady=6)

        # ── Right: options + action ───────────────────────────────────────────
        self._section_label(right, "OPTIONS")
        opts_frame = ctk.CTkFrame(right, fg_color=BG_FIELD, corner_radius=8)
        opts_frame.pack(fill="x", pady=(4, 0))

        self._dry_run   = tk.BooleanVar(value=True)
        self._backup    = tk.BooleanVar(value=True)
        self._both_exts = tk.BooleanVar(value=True)

        ctk.CTkCheckBox(opts_frame, text="Dry run (preview only)",
                        variable=self._dry_run,
                        font=("Courier New", 11), text_color=TEXT_MAIN,
                        fg_color=ACCENT, hover_color=ACCENT2).pack(
                            anchor="w", padx=14, pady=8)
        ctk.CTkCheckBox(opts_frame, text="Backup originals",
                        variable=self._backup,
                        font=("Courier New", 11), text_color=TEXT_MAIN,
                        fg_color=ACCENT, hover_color=ACCENT2).pack(
                            anchor="w", padx=14, pady=6)
        ctk.CTkCheckBox(opts_frame, text="Rename .jpg + .png pairs",
                        variable=self._both_exts,
                        font=("Courier New", 11), text_color=TEXT_MAIN,
                        fg_color=ACCENT, hover_color=ACCENT2).pack(
                            anchor="w", padx=14, pady=(6, 12))

        ctk.CTkButton(right, text="⟳  Preview Changes", height=38,
                      fg_color=BG_FIELD, border_width=1, border_color=ACCENT,
                      hover_color=BG_MID, font=("Courier New", 12, "bold"),
                      text_color=ACCENT,
                      command=self._preview_rename).pack(fill="x", pady=(18, 6))
        ctk.CTkButton(right, text="✔  Apply Rename", height=44,
                      fg_color=ACCENT, hover_color=ACCENT2,
                      font=("Courier New", 13, "bold"),
                      text_color="white",
                      command=self._apply_rename).pack(fill="x", pady=6)

        # ── Log ───────────────────────────────────────────────────────────────
        self._section_label(left, "OPERATION LOG")
        self._rename_log = ctk.CTkTextbox(left, height=160,
                                          font=("Courier New", 10),
                                          fg_color=BG_FIELD, text_color=TEXT_MAIN,
                                          border_color=BORDER, corner_radius=6)
        self._rename_log.pack(fill="both", expand=True, pady=(4, 0))

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
            ("Seq From", "f_seq_s",  None),
            ("Seq To",   "f_seq_e",  None),
        ]
        self._filter_vars: dict[str, tk.StringVar] = {}
        cols = 4
        for idx, (label, key, opts) in enumerate(filter_defs):
            r, c = divmod(idx, cols)
            ctk.CTkLabel(crit, text=label, font=("Courier New", 10),
                         text_color=TEXT_DIM, anchor="w").grid(
                             row=r*2, column=c, sticky="w", padx=10, pady=(8, 0))
            var = tk.StringVar(value="")
            self._filter_vars[key] = var
            if opts:
                mb = ctk.CTkOptionMenu(crit, variable=var,
                                       values=["(any)"] + opts,
                                       fg_color=BG_CARD, button_color=ACCENT2,
                                       button_hover_color=ACCENT,
                                       text_color=TEXT_MAIN,
                                       font=("Courier New", 11), width=140)
                mb.set("(any)")
                mb.grid(row=r*2+1, column=c, sticky="w", padx=10, pady=(0, 8))
            else:
                ctk.CTkEntry(crit, textvariable=var, width=120,
                             font=("Courier New", 11), fg_color=BG_CARD,
                             text_color=TEXT_MAIN, border_color=BORDER,
                             placeholder_text="—").grid(
                                 row=r*2+1, column=c, sticky="w",
                                 padx=10, pady=(0, 8))

        # extension filter
        self._filter_ext = tk.StringVar(value="both")
        ext_row = ctk.CTkFrame(top, fg_color="transparent")
        ext_row.pack(fill="x", pady=(8, 0))
        ctk.CTkLabel(ext_row, text="Type:", font=("Courier New", 10),
                     text_color=TEXT_DIM).pack(side="left", padx=(0, 8))
        for val, txt in [("both", "Color + Depth"), ("jpg", "Color only (.jpg)"),
                          ("png", "Depth only (.png)")]:
            ctk.CTkRadioButton(ext_row, text=txt, variable=self._filter_ext,
                               value=val, font=("Courier New", 11),
                               text_color=TEXT_MAIN,
                               fg_color=ACCENT, hover_color=ACCENT2).pack(
                                   side="left", padx=12)

        btn_row = ctk.CTkFrame(top, fg_color="transparent")
        btn_row.pack(fill="x", pady=(10, 0))
        ctk.CTkButton(btn_row, text="⊞  Search", height=38, width=140,
                      fg_color=ACCENT, hover_color=ACCENT2,
                      font=("Courier New", 12, "bold"),
                      command=self._run_filter).pack(side="left")
        self._filter_count = tk.StringVar(value="")
        ctk.CTkLabel(btn_row, textvariable=self._filter_count,
                     font=("Courier New", 11), text_color=SUCCESS).pack(
                         side="left", padx=16)
        ctk.CTkButton(btn_row, text="⊡ Export list", height=38, width=130,
                      fg_color=BG_FIELD, border_width=1, border_color=ACCENT,
                      hover_color=BG_MID, font=("Courier New", 11),
                      text_color=ACCENT,
                      command=self._export_filter_list).pack(side="left", padx=8)
        ctk.CTkButton(btn_row, text="⊠ Copy matched files", height=38, width=170,
                      fg_color=BG_FIELD, border_width=1, border_color=ACCENT2,
                      hover_color=BG_MID, font=("Courier New", 11),
                      text_color=ACCENT2,
                      command=self._copy_matched).pack(side="left", padx=4)

        # results
        self._section_label(parent, "RESULTS", padx=12)
        cols_frame = ctk.CTkFrame(parent, fg_color=BG_FIELD, corner_radius=0, height=26)
        cols_frame.pack(fill="x", padx=12)
        cols_frame.pack_propagate(False)
        for txt, w in [("Filename", 38), ("Room", 8), ("Height", 7),
                        ("Angle", 8), ("Distance", 9), ("Lighting", 8),
                        ("Seq", 6), ("Ext", 5)]:
            ctk.CTkLabel(cols_frame, text=txt.upper(), font=("Courier New", 9, "bold"),
                         text_color=ACCENT, width=w*7, anchor="w").pack(side="left", padx=4)

        self._filter_results = ctk.CTkTextbox(parent, font=("Courier New", 10),
                                              fg_color=BG_FIELD,
                                              text_color=TEXT_MAIN,
                                              border_color=BORDER, corner_radius=0)
        self._filter_results.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self._filter_matches: list[tuple[str, str, dict]] = []   # (abs, rel, parts)

    # ══════════════════════════════════════════════════════════════════════════
    #  TAB 3 — PREVIEW
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
                     font=("Courier New", 11), text_color=SUCCESS).pack(
                         side="left", padx=16)

        # columns header
        hdr = ctk.CTkFrame(parent, fg_color=BG_MID, corner_radius=0, height=24)
        hdr.pack(fill="x", padx=12)
        hdr.pack_propagate(False)
        for txt, w in [("Path", 35), ("Room", 8), ("Height", 7),
                        ("Angle", 10), ("Distance", 9), ("Lighting", 9),
                        ("Seq", 6), ("Type", 6)]:
            ctk.CTkLabel(hdr, text=txt, font=("Courier New", 9, "bold"),
                         text_color=ACCENT, width=w*7, anchor="w").pack(
                             side="left", padx=4)

        self._preview_box = ctk.CTkTextbox(parent, font=("Courier New", 10),
                                           fg_color=BG_FIELD, text_color=TEXT_MAIN,
                                           border_color=BORDER, corner_radius=0)
        self._preview_box.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _section_label(self, parent, text, padx=0):
        ctk.CTkLabel(parent, text=text, font=("Courier New", 9, "bold"),
                     text_color=ACCENT).pack(anchor="w", padx=padx, pady=(8, 2))

    def _log(self, widget, text: str, tag: str = ""):
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
    def _gather_rename_plan(self, dry: bool = True) -> list[tuple[str, str]]:
        """Return list of (old_abs_path, new_abs_path) pairs."""
        root = self._get_root()
        if not root:
            return []

        rv    = self._rename_vars
        rfv   = self._rfilter_vars
        seq_s = self._seq_start.get().strip()
        seq_e = self._seq_end.get().strip()

        plan: list[tuple[str, str]] = []
        seen_new: set[str] = set()

        for abs_path, rel_folder, fname in walk_images(root):
            parts = parse_filename(fname)
            if not parts:
                continue

            # ── apply filter (only rename matching files)
            if rfv["rf_room"].get()   and rfv["rf_room"].get()   != "(any)" \
               and parts["room"]    != rfv["rf_room"].get():   continue
            if rfv["rf_height"].get() and rfv["rf_height"].get() != "(any)" \
               and parts["height"]  != rfv["rf_height"].get(): continue
            if rfv["rf_angle"].get()  and rfv["rf_angle"].get()  != "(any)" \
               and parts["angle"]   != rfv["rf_angle"].get():  continue
            if rfv["rf_dist"].get()   and rfv["rf_dist"].get()   != "(any)" \
               and parts["distance"]!= rfv["rf_dist"].get():   continue
            if rfv["rf_light"].get()  and rfv["rf_light"].get()  != "(any)" \
               and parts["lighting"] != rfv["rf_light"].get():  continue

            # ── apply sequence range
            if seq_s and int(parts["sequence"]) < int(seq_s): continue
            if seq_e and int(parts["sequence"]) > int(seq_e): continue

            # ── extension filter (both_exts option does nothing for filter)
            new = dict(parts)
            def apply(key, new_key):
                val = rv[new_key].get()
                if val and val != "(keep)":
                    new[key] = val
            apply("room",     "new_room")
            apply("height",   "new_height")
            apply("angle",    "new_angle")
            apply("distance", "new_dist")
            apply("lighting", "new_light")

            new_name     = build_filename(new)
            new_abs      = os.path.join(os.path.dirname(abs_path), new_name)
            if new_name == fname:
                continue
            if new_abs in seen_new:
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
            old_n = os.path.basename(old)
            new_n = os.path.basename(new)
            self._log(self._rename_log, f"  {old_n:<53}  →  {new_n}")
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
        root    = self.dataset_path.get().strip()
        backup  = self._backup.get()
        errors  = 0
        renamed = 0

        for old, new in plan:
            try:
                if backup:
                    bu = old + ".bak"
                    shutil.copy2(old, bu)
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
    #  FILTER / SEARCH LOGIC
    # ══════════════════════════════════════════════════════════════════════════
    def _run_filter(self):
        root = self._get_root()
        if not root:
            return

        fv    = self._filter_vars
        ext   = self._filter_ext.get()
        seq_s = fv["f_seq_s"].get().strip()
        seq_e = fv["f_seq_e"].get().strip()

        def match(val, key, opt_key=None):
            chosen = fv[key].get().strip()
            if not chosen or chosen == "(any)":
                return True
            return val == chosen

        self._filter_matches = []
        self._clear_log(self._filter_results)

        for abs_path, rel_folder, fname in walk_images(root):
            parts = parse_filename(fname)
            if not parts:
                continue

            # extension
            if ext == "jpg" and parts["ext"] != ".jpg": continue
            if ext == "png" and parts["ext"] != ".png": continue

            if not match(parts["room"],     "f_room"):    continue
            if not match(parts["height"],   "f_height"):  continue
            if not match(parts["angle"],    "f_angle"):   continue
            if not match(parts["distance"], "f_dist"):    continue
            if not match(parts["lighting"], "f_light"):   continue

            if seq_s and int(parts["sequence"]) < int(seq_s): continue
            if seq_e and int(parts["sequence"]) > int(seq_e): continue

            self._filter_matches.append((abs_path, rel_folder, parts))

        total = len(self._filter_matches)
        self._filter_count.set(f"{total} match(es)")

        for abs_path, rel_folder, parts in self._filter_matches:
            row = (f"  {parts['original']:<42}  "
                   f"{parts['room']:8}  "
                   f"{parts['height']:7}  "
                   f"{ANGLE_LABEL.get(parts['angle'], parts['angle']):10}  "
                   f"{parts['distance']:9}  "
                   f"{LIGHT_LABEL.get(parts['lighting'], parts['lighting']):8}  "
                   f"{parts['sequence']:6}  "
                   f"{parts['ext']}")
            self._log(self._filter_results, row)

        self._set_status(f"Filter: {total} result(s)")

    def _export_filter_list(self):
        if not self._filter_matches:
            messagebox.showinfo("Empty", "Run a search first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text file", "*.txt"), ("All", "*.*")],
            title="Export list to…")
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write("abs_path\trelative_folder\troom\theight\tangle\tdistance"
                    "\tlighting\tsequence\text\n")
            for abs_p, rel_f, parts in self._filter_matches:
                f.write(f"{abs_p}\t{rel_f}\t{parts['room']}\t{parts['height']}\t"
                        f"{parts['angle']}\t{parts['distance']}\t{parts['lighting']}\t"
                        f"{parts['sequence']}\t{parts['ext']}\n")
        messagebox.showinfo("Exported", f"Saved {len(self._filter_matches)} rows to\n{path}")

    def _copy_matched(self):
        if not self._filter_matches:
            messagebox.showinfo("Empty", "Run a search first.")
            return
        dest = filedialog.askdirectory(title="Copy matched files to…")
        if not dest:
            return
        copied = 0
        for abs_p, _, _ in self._filter_matches:
            try:
                shutil.copy2(abs_p, dest)
                copied += 1
            except Exception as e:
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
        total = 0
        bad   = 0
        for abs_path, rel_folder, fname in walk_images(root):
            parts = parse_filename(fname)
            if parts:
                ftype = "color" if parts["ext"] == ".jpg" else "depth"
                row   = (f"  {rel_folder:<30}  "
                         f"{parts['room']:8}  "
                         f"{parts['height']:7}  "
                         f"{ANGLE_LABEL.get(parts['angle'],parts['angle']):10}  "
                         f"{parts['distance']:9}  "
                         f"{LIGHT_LABEL.get(parts['lighting'],parts['lighting']):8}  "
                         f"{parts['sequence']:6}  "
                         f"{ftype}")
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
