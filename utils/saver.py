import os
import re
import cv2
import numpy as np


def get_save_path(root_dir: str, dataset: str, floor_num: str, room_num: str) -> str:
    return os.path.join(root_dir, dataset, floor_num.strip(), room_num.strip())


def get_next_sequence(save_path: str) -> str:
    """Scans existing color saves and returns the next zero-padded sequence number."""
    color_dir = os.path.join(save_path, "color")

    if not os.path.exists(color_dir):
        return "0001"

    max_seq = 0
    for fname in os.listdir(color_dir):
        if fname.endswith(".jpg"):
            match = re.search(r'_(\d{4})\.jpg$', fname)
            if match:
                num = int(match.group(1))
                if num > max_seq:
                    max_seq = num

    return str(max_seq + 1).zfill(4)


def save_frame(
    save_path: str,
    color_image: np.ndarray,
    depth_image: np.ndarray,
    ffrrrr: str,
    height: str,
    angle: str,
    distance_category: str,
    lighting: str,
    sequence: str,
) -> tuple[str, str]:
    """
    Saves a color JPEG and a 16-bit depth PNG.

    Returns:
        (color_filepath, depth_filepath)
    """
    color_dir = os.path.join(save_path, "color")
    depth_dir = os.path.join(save_path, "depth_raw")
    os.makedirs(color_dir, exist_ok=True)
    os.makedirs(depth_dir, exist_ok=True)

    base_name  = f"{ffrrrr}_{height}_{angle}_{distance_category}_{lighting}_{sequence}"
    color_file = os.path.join(color_dir, f"{base_name}.jpg")
    depth_file = os.path.join(depth_dir, f"{base_name}_depth.png")

    cv2.imwrite(color_file, color_image)
    cv2.imwrite(depth_file, depth_image)

    return color_file, depth_file
