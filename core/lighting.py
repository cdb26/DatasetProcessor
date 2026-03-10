"""
core/lighting.py
Lighting detection based on image brightness.
"""

import cv2
import numpy as np

DIM_THRESHOLD  = 80
WELL_THRESHOLD = 120


def detect_lighting(color_image: np.ndarray, previous: str = "well") -> tuple[str, float]:
    """
    Detects scene lighting from a BGR image.

    Args:
        color_image: BGR numpy array
        previous: previous lighting label (prevents flicker in mid-range)

    Returns:
        lighting (str): 'well' | 'dim'
        brightness (float): mean grayscale value 0–255
    """
    gray = cv2.cvtColor(color_image, cv2.COLOR_BGR2GRAY)
    brightness = float(np.mean(gray))

    if brightness < DIM_THRESHOLD:
        lighting = "dim"
    elif brightness > WELL_THRESHOLD:
        lighting = "well"
    else:
        lighting = previous  # hysteresis – avoid flickering

    return lighting, brightness
