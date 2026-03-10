import cv2
import numpy as np


def draw_crosshair(image: np.ndarray, x: int, y: int, size: int = 15, color=(0, 255, 80), thickness: int = 2):
    cv2.line(image, (x - size, y), (x + size, y), color, thickness)
    cv2.line(image, (x, y - size), (x, y + size), color, thickness)


def draw_distance_text(image: np.ndarray, dist: float, x: int, y: int):
    cv2.putText(
        image, f"{dist:.2f}m",
        (x + 18, y - 12),
        cv2.FONT_HERSHEY_SIMPLEX, 0.65, (0, 255, 80), 2,
    )


def draw_lighting_text(image: np.ndarray, lighting: str, brightness: float):
    cv2.putText(
        image, f"Light: {lighting}  ({brightness:.0f})",
        (12, 32),
        cv2.FONT_HERSHEY_SIMPLEX, 0.65, (0, 220, 255), 2,
    )


def annotate_frame(
    color_image: np.ndarray,
    dist: float,
    cx: int,
    cy: int,
    lighting: str,
    brightness: float,
) -> np.ndarray:
    """Applies all overlays and returns a flipped display-ready image."""
    display = color_image.copy()
    display = cv2.flip(display, 1)
    draw_crosshair(display, cx, cy)
    draw_distance_text(display, dist, cx, cy)
    draw_lighting_text(display, lighting, brightness)
    return display
