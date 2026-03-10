"""
core/camera.py
Handles RealSense pipeline initialization, frame acquisition, and alignment.
"""

import pyrealsense2 as rs
import numpy as np


class RealSenseCamera:
    MIN_DIST = 0.1
    MAX_DIST = 10.0

    def __init__(self, width=1280, height=720, fps=30):
        self.width = width
        self.height = height
        self.fps = fps

        self.pipeline = rs.pipeline()
        self.config = rs.config()
        self.align = None
        self.depth_scale = None
        self.profile = None

    def start(self):
        self.config.enable_stream(rs.stream.color, self.width, self.height, rs.format.bgr8, self.fps)
        self.config.enable_stream(rs.stream.depth, self.width, self.height, rs.format.z16, self.fps)

        self.profile = self.pipeline.start(self.config)
        self.align = rs.align(rs.stream.color)

        depth_sensor = self.profile.get_device().first_depth_sensor()
        self.depth_scale = depth_sensor.get_depth_scale()

    def get_frames(self):
        """
        Returns (color_image: np.ndarray, depth_image: np.ndarray)
        or (None, None) if frames are unavailable.
        """
        frames = self.pipeline.wait_for_frames()
        aligned = self.align.process(frames)

        color_frame = aligned.get_color_frame()
        depth_frame = aligned.get_depth_frame()

        if not color_frame or not depth_frame:
            return None, None

        color_image = np.asanyarray(color_frame.get_data())
        depth_image = np.asanyarray(depth_frame.get_data())

        return color_image, depth_image

    def stop(self):
        self.pipeline.stop()
