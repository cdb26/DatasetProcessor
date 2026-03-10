# Depth Dataset Capture

A modular desktop application for capturing paired **colour + depth** images
using an Intel RealSense camera.

---

## Project Structure

```
depth_dataset_app/
│
├── main.py                   # Entry-point; assembles and runs the App
│
├── core/                     # Hardware & perception logic (no UI)
│   ├── camera.py             # RealSense pipeline wrapper
│   ├── depth.py              # Distance tracking & categorisation
│   └── lighting.py           # Lighting detection (brightness-based)
│
├── ui/                       # GUI layer
│   ├── theme.py              # Palette, fonts, size constants
│   └── widgets.py            # Reusable themed widget factories
│
├── utils/                    # Stateless helpers
│   ├── overlay.py            # OpenCV annotation helpers
│   └── saver.py              # File naming, directory creation, imwrite
│
└── requirements.txt
```

---

## Installation

```bash
pip install -r requirements.txt
```

> `pyrealsense2` may require the [Intel RealSense SDK](https://github.com/IntelRealSense/librealsense/releases).

---

## Running

```bash
python main.py
```

---

## Usage

| Control | Action |
|---|---|
| **Enable Camera Mode** button | Locks editing fields, starts live capture mode |
| **P** key | Captures a frame (camera mode must be active) |
| **Capture [P]** button | Same as pressing P |
| **Disable Camera Mode** button | Returns to editing mode |

### File Naming

Saved files follow the pattern:

```
{FFRRRR}_{height}_{distance_category}_{lighting}_{sequence}
```

Example: `010203_1.2m_medium_well_0007.jpg`

### Distance Categories

| Range | Category |
|---|---|
| ≤ 1.0 m | `close` |
| 1.0 – 1.6 m | `medium` |
| > 1.6 m | `far` |

---

## Dataset Layout

```
dataset/
└── <floorNum>/
    └── <roomNum>/
        ├── color/
        │   └── *.jpg
        └── depth_raw/
            └── *_depth.png
```
