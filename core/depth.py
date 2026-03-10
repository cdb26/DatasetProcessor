"""
core/depth.py
Depth analysis: distance tracking and distance category classification.
"""

import numpy as np


MAX_DIST = 10.0


def track_distance(depth_meters: np.ndarray, patch_size: int = 5):
    """
    Computes the median depth in a small central region.

    Returns:
        dist (float): median distance in meters
        category (str): 'close' | 'medium' | 'far'
        cx (int): center x pixel
        cy (int): center y pixel
    """
    h, w = depth_meters.shape
    cx, cy = w // 2, h // 2

    region = depth_meters[cy - patch_size: cy + patch_size,
                           cx - patch_size: cx + patch_size]
    valid = region[region > 0]

    dist = float(np.median(valid)) if len(valid) > 0 else MAX_DIST
    category = classify_distance(dist)

    return dist, category, cx, cy


def classify_distance(dist: float) -> str:
    """Maps a distance in meters to a named category."""
    if dist <= 1.0:
        return "close"
    elif dist <= 1.6:
        return "medium"
    return "far"
