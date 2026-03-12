
import os
import cv2
import customtkinter as ctk
from PIL import Image, ImageTk

from core.camera   import RealSenseCamera
from core.depth    import track_distance
from core.lighting import detect_lighting
from utils.overlay import annotate_frame
from utils.saver   import get_save_path, get_next_sequence, save_frame
from ui.theme      import *
from ui.widgets    import (
    make_label, make_value_label, make_entry,
    make_option_menu, make_button, make_card, section_title,
)

class App(ctk.CTk):
    DATASET   = "dataset"
    ROOT_DIR  = os.getcwd()

    def __init__(self):
        super().__init__()
        self._configure_window()
        self._init_state()
        self._build_ui()

        self.camera = RealSenseCamera()
        self.camera.start()

        self._update_frame()


    def _configure_window(self):
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.title("Dataset Capture")
        self.geometry("1460x860")
        self.configure(fg_color=BG_DARK)
        self.resizable(False, False)

    def _init_state(self):
        self.camera_mode       = False
        self.distance_category = "unknown"
        self.latest_color      = None
        self.latest_depth      = None
        
        self.ffrrrr_var   = ctk.StringVar(value="FFRRRR")
        self.height_var   = ctk.StringVar(value="0.8m")
        self.angle_var = ctk.StringVar(value="1")
        self.lighting_var = ctk.StringVar(value="well")
        self.sequence_var = ctk.StringVar(value="0001")
        self.floor_var    = ctk.StringVar(value="floorNum")
        self.room_var     = ctk.StringVar(value="roomNum")

        self.floor_var.trace_add("write", lambda *_: self._refresh_path())
        self.room_var.trace_add("write",  lambda *_: self._refresh_path())

    def _build_ui(self):
        self._build_header()

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        body.columnconfigure(0, weight=5)
        body.columnconfigure(2, weight=2)
        body.rowconfigure(0, weight=2)

        self._build_camera_panel(body)
        self._build_side_panel(body)

        self.bind("<Key>", self._key_handler)

    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color="transparent", height=56)
        header.pack(fill="x", padx=20, pady=(14, 8))
        header.pack_propagate(False)

        title = ctk.CTkLabel(
            header,
            text="DEPTH  DATASET  CAPTURE",
            font=("Syne", 18, "bold"),
            text_color=TEXT_PRIMARY,
        )
        title.pack(side="left", anchor="w")

        dot = ctk.CTkLabel(header, text="●", font=("Syne", 12), text_color=TEXT_MUTED)
        dot.pack(side="left", padx=(14, 6), anchor="w")

        self.mode_badge = ctk.CTkLabel(
            header,
            text="EDITING",
            font=("Syne", 11, "bold"),
            text_color=TEXT_MUTED,
        )
        self.mode_badge.pack(side="left", anchor="w")

        self.toggle_btn = make_button(
            header,
            text="Enable Camera Mode",
            command=self._toggle_camera_mode,
            width=180,
        )
        self.toggle_btn.pack(side="right", anchor="e")

    def _build_camera_panel(self, parent):
        card = make_card(parent)
        card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self.image_label = ctk.CTkLabel(card, text="")
        self.image_label.pack(expand=True, fill="both", padx=4, pady=4)

    def _build_side_panel(self, parent):
        panel = ctk.CTkFrame(parent, fg_color="transparent")
        panel.grid(row=0, column=1, sticky="nsew")

        self._build_status_cards(panel)
        self._build_settings_card(panel)
        self._build_path_card(panel)
        self._build_capture_card(panel)

    def _build_status_cards(self, parent):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=(0, 10))
        row.columnconfigure((0, 1), weight=1)

        d_card = make_card(row)
        d_card.grid(row=0, column=0, sticky="ew", padx=(0, 5), pady=0)
        make_label(d_card, "DISTANCE").pack(anchor="w", padx=14, pady=(12, 0))
        self.dist_value = make_value_label(d_card, "--")
        self.dist_value.pack(anchor="w", padx=14, pady=(0, 4))
        self.dist_cat = make_label(d_card, "category: --", color=TEXT_MUTED)
        self.dist_cat.pack(anchor="w", padx=14, pady=(0, 12))

        l_card = make_card(row)
        l_card.grid(row=0, column=1, sticky="ew", padx=(5, 0))
        make_label(l_card, "LIGHTING").pack(anchor="w", padx=14, pady=(12, 0))
        self.light_value = make_value_label(l_card, "--")
        self.light_value.pack(anchor="w", padx=14, pady=(0, 4))
        self.bright_label = make_label(l_card, "brightness: --", color=TEXT_MUTED)
        self.bright_label.pack(anchor="w", padx=14, pady=(0, 12))

    def _build_settings_card(self, parent):
        card = make_card(parent)
        card.pack(fill="x", pady=(0, 10))

        section_title(card, "CAPTURE SETTINGS").pack(anchor="w", padx=14, pady=(12, 8))

        grid = ctk.CTkFrame(card, fg_color="transparent")
        grid.pack(fill="x", padx=14, pady=(0, 14))
        grid.columnconfigure((0, 1), weight=1)

        make_label(grid, "FFRRRR").grid(row=0, column=0, sticky="w", pady=(0, 2))
        self.ffrrrr_entry = make_entry(grid, self.ffrrrr_var)
        self.ffrrrr_entry.grid(row=1, column=0, sticky="ew", padx=(0, 6), pady=(0, 10))

        make_label(grid, "SEQUENCE").grid(row=0, column=1, sticky="w", pady=(0, 2))
        self.seq_entry = make_entry(grid, self.sequence_var, width=80)
        self.seq_entry.grid(row=1, column=1, sticky="ew", padx=(6, 0), pady=(0, 10))
        

        make_label(grid, "HEIGHT").grid(row=2, column=0, sticky="w", pady=(0, 2))
        self.height_menu = make_option_menu(grid, ["0.8m", "1.2m", "1.6m"], self.height_var)
        self.height_menu.grid(row=3, column=0, sticky="ew", padx=(0, 6))
        
        make_label(grid, "ANGLE").grid(row=4, column=0, sticky="w", pady=(0,2))
        self.angle_menu = make_option_menu(grid, ["1", "2", "3"], self.angle_var)
        self.angle_menu.grid(row=5, column=0, sticky="ew")
        make_label(grid, "1 - Ortho 2 - Diagonal 3 - Top Down").grid(row=4, column=1, sticky="ew")

        make_label(grid, "LIGHTING OVERRIDE").grid(row=2, column=1, sticky="w", pady=(0, 2))
        self.lighting_menu = make_option_menu(grid, ["well", "dim"], self.lighting_var)
        self.lighting_menu.grid(row=3, column=1, sticky="ew", padx=(6, 0))

        self._settings_widgets = [
            self.ffrrrr_entry, self.seq_entry,
            self.height_menu, self.lighting_menu,
        ]

    def _build_path_card(self, parent):
        card = make_card(parent)
        card.pack(fill="x", pady=(0, 10))

        section_title(card, "SAVE PATH").pack(anchor="w", padx=14, pady=(12, 8))

        grid = ctk.CTkFrame(card, fg_color="transparent")
        grid.pack(fill="x", padx=14, pady=(0, 10))
        grid.columnconfigure((0, 1), weight=1)

        make_label(grid, "FLOOR").grid(row=0, column=0, sticky="w", pady=(0, 2))
        self.floor_entry = make_entry(grid, self.floor_var)
        self.floor_entry.grid(row=1, column=0, sticky="ew", padx=(0, 6))

        make_label(grid, "ROOM").grid(row=0, column=1, sticky="w", pady=(0, 2))
        self.room_entry = make_entry(grid, self.room_var)
        self.room_entry.grid(row=1, column=1, sticky="ew", padx=(6, 0))

        self.path_label = ctk.CTkLabel(
            card,
            text="",
            font=FONT_MONO,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
        )
        self.path_label.pack(fill="x", padx=14, pady=(8, 14))

        self._settings_widgets += [self.floor_entry, self.room_entry]
        self._refresh_path()

    def _build_capture_card(self, parent):
        card = make_card(parent)
        card.pack(fill="x")

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=14)

        self.capture_btn = make_button(
            inner,
            text="Capture  [P]",
            command=self._capture_image,
            width=220,
            color=ACCENT,
            hover=ACCENT_DIM,
        )
        
        self.switch_btn = make_button(
            inner,
            text="Switch Angle [K]",
            command=self._switch_angle,
            width=220,
            color=ACCENT,
            hover=ACCENT_DIM,
        )
        
        self.switch_btn.pack(fill="x", pady=(0,6))
        self.capture_btn.pack(fill="x")

        self.capture_status = make_label(inner, "", color=TEXT_MUTED)
        self.capture_status.pack(pady=(6, 0))
        
    def _switch_angle(self):
        angles = ["1", "2", "3"]
        current = self.angle_var.get()

        try:
            idx = angles.index(current)
            next_angle = angles[(idx + 1) % len(angles)]
        except ValueError:
            next_angle = "1"

        self.angle_var.set(next_angle)

    def _toggle_camera_mode(self):
        self.camera_mode = not self.camera_mode

        if self.camera_mode:
            for w in self._settings_widgets:
                w.configure(state="disabled")
            self.sequence_var.set(get_next_sequence(self._get_save_path()))
            self.toggle_btn.configure(
                text="Disable Camera Mode",
                fg_color=DANGER,
                hover_color="#CC3344",
            )
            self.mode_badge.configure(text="LIVE CAPTURE", text_color=ACCENT)
            self.focus_set()
        else:
            for w in self._settings_widgets:
                w.configure(state="normal")
            self.toggle_btn.configure(
                text="Enable Camera Mode",
                fg_color=ACCENT,
                hover_color=ACCENT_DIM,
            )
            self.mode_badge.configure(text="EDITING", text_color=TEXT_MUTED)

    def _get_save_path(self):
        return get_save_path(
            self.ROOT_DIR, self.DATASET,
            self.floor_var.get(), self.room_var.get(),
        )

    def _refresh_path(self):
        path = self._get_save_path()
        rel  = os.path.relpath(path, self.ROOT_DIR)
        self.path_label.configure(
            text=f"./{rel}/\n  ├── color/\n  └── depth_raw/"
        )
        

    def _capture_image(self, event=None):
        if not self.camera_mode or self.latest_color is None:
            return

        color_file, depth_file = save_frame(
            save_path=self._get_save_path(),
            color_image=self.latest_color,
            depth_image=self.latest_depth,
            ffrrrr=self.ffrrrr_var.get(),
            height=self.height_var.get(),
            angle=self.angle_var.get(),
            distance_category=self.distance_category,
            lighting=self.lighting_var.get(),
            sequence=self.sequence_var.get(),
        )

        print(f"Saved: {color_file}")
        print(f"Saved: {depth_file}")

        seq = int(self.sequence_var.get())
        self.sequence_var.set(str(seq + 1).zfill(4))

        name = os.path.basename(color_file)
        self.capture_status.configure(text=f"✓  {name}", text_color=ACCENT)
        self.after(2000, lambda: self.capture_status.configure(text=""))
        

    def _key_handler(self, event):
        if self.camera_mode and event.char.lower() == "p":
            self._capture_image()

        if self.camera_mode and event.char.lower() == "k":
            self._switch_angle()
                

    def _update_frame(self):
        color_image, depth_image = self.camera.get_frames()

        if color_image is None:
            self.after(10, self._update_frame)
            return
        
        floor = self.floor_var.get()
        room = self.room_var.get()

        if floor != "floorNum" and room != "roomNum":
            self.ffrrrr_var.set(f"{floor}{room}")
        
        lighting, brightness = detect_lighting(color_image, self.lighting_var.get())
        self.lighting_var.set(lighting)

        self.latest_color = color_image.copy()
        self.latest_depth = depth_image
        
        depth_meters = depth_image * self.camera.depth_scale
        dist, category, cx, cy = track_distance(depth_meters)
        self.distance_category = category

        self.dist_value.configure(text=f"{dist:.2f} m")
        self.dist_cat.configure(text=f"category: {category}")
        self.light_value.configure(
            text=lighting,
            text_color=WARNING if lighting == "dim" else ACCENT,
        )
        self.bright_label.configure(text=f"brightness: {brightness:.0f}")

        display = annotate_frame(color_image, dist, cx, cy, lighting, brightness)
        img_rgb = cv2.cvtColor(display, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(img_rgb)

        pil_img = pil_img.resize((960, 540), Image.LANCZOS)
        imgtk = ImageTk.PhotoImage(pil_img)
        self.image_label.imgtk = imgtk
        self.image_label.configure(image=imgtk)

        self.after(30, self._update_frame)

    def destroy(self):
        self.camera.stop()
        super().destroy()   

if __name__ == "__main__":
    app = App()
    app.mainloop()
