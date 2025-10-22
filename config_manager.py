"""
config_manager.py
-----------------
Manages configuration, paths, and environment setup
for the MDA Camera Automation system.

Author: Syrine Bouhoula / MDA Project
Date: 2025
"""

import os
import json
import time


class ConfigManager:
    """Handles configuration loading and directory setup."""

    def __init__(self, config_path="camera_config.json"):
        self.config_path = config_path
        self.config = self.load_config()
        self._setup_environment()
        self._extract_values()

    # ====== Load configuration ======
    def load_config(self):
        """Load camera and category configuration from JSON file."""
        try:
            with open(self.config_path, "r") as f:
                cfg = json.load(f)
            print(f" Loaded configuration from {self.config_path}")
            return cfg
        except FileNotFoundError:
            print(f"Config file '{self.config_path}' not found. Using defaults.")
            return {}
        except Exception as e:
            print(f" Error reading config: {e}")
            return {}

    # ====== Setup paths & environment ======
    def _setup_environment(self):
        """Create working directories and register DLL paths."""
        self.HERE = os.path.dirname(os.path.abspath(__file__))
        self.timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
        self.PROJECT_ID = self.config.get("project", {}).get("PROJECT_ID", "MDA12345")
        self.DEVICE_CODE = self.config.get("project", {}).get("DEVICE_CODE", "DC123")

        # Session folder for this run
        self.session_dir = os.path.join(self.HERE, f"{self.PROJECT_ID}_{self.timestamp}")
        os.makedirs(self.session_dir, exist_ok=True)
        self.CAPTURE_DIR = self.session_dir
        self.LOG_PATH = os.path.join(
            self.session_dir,
            f"{self.PROJECT_ID}_{time.strftime('%Y-%m-%d')}_{self.DEVICE_CODE}.xlsx"
        )

        # Add Canon EDSDK DLL search paths on Windows
        dll_paths = [
            os.path.join(self.HERE, "EDSDK_64", "Dll"),
            os.path.join(self.HERE, "EDSDK", "Dll"),
            os.path.join(self.HERE, "..", "EDSDK_64", "Dll"),
            os.path.join(self.HERE, "..", "EDSDK", "Dll"),
        ]
        for p in dll_paths:
            if hasattr(os, "add_dll_directory") and os.path.isdir(p):
                try:
                    os.add_dll_directory(os.path.abspath(p))
                except Exception:
                    pass

    # ====== Extract main parameters ======
    def _extract_values(self):
        """Extract useful config fields for easy access."""
        cam = self.config.get("camera", {})
        self.AV_LABEL = cam.get("AV_LABEL", "f/8")
        self.ISO_LABEL = cam.get("ISO_LABEL", "100")
        self.TV_REF_LABEL = cam.get("TV_REF_LABEL", "1/60")
        self.DELAY_S = cam.get("DELAY_S", 3.0)
        self.POST_SHOT_WAIT = cam.get("POST_SHOT_WAIT", 2.5)
        self.THUMBNAIL_WIDTH_PX = cam.get("THUMBNAIL_WIDTH_PX", 180)

        zoom_cfg = cam.get("ZOOM_STEPS", {})
        self.ZOOM_140_STR = zoom_cfg.get("ZOOM_140_STR", "140")
        self.ZOOM_120_STR = zoom_cfg.get("ZOOM_120_STR", "120")
        self.ZOOM_110_STR = zoom_cfg.get("ZOOM_110_STR", "110")
        self.ZOOM_100_STR = zoom_cfg.get("ZOOM_100_STR", "100")
        self.ZOOM_055_STR = zoom_cfg.get("ZOOM_055_STR", "55")

        cats = self.config.get("categories", {})
        self.CAT_REF = cats.get("CAT_REF", "reference focus sticker")

        orders = self.config.get("orders", {})
        self.FEATURE_ORDER = [cats.get(k, k) for k in orders.get("FEATURE_ORDER", [])]
        self.ORDER_Z140 = [cats.get(k, k) for k in orders.get("ORDER_Z140", [])]
        self.ORDER_Z120 = [cats.get(k, k) for k in orders.get("ORDER_Z120", [])]
        self.ORDER_Z055 = [cats.get(k, k) for k in orders.get("ORDER_Z055", [])]
        self.ORDER_Z110 = [cats.get(k, k) for k in orders.get("ORDER_Z110", [])]

        tv_map_cfg = self.config.get("tv_map", {})
        self.TV_MAP = {cats.get(k, k): v for k, v in tv_map_cfg.items()}

        print(f"📁 Session directory: {self.session_dir}")
        print(f"🧾 Log path: {self.LOG_PATH}")

    # ====== Public helper methods ======
    def get_project_info(self):
        return self.config.get("project", {})

    def get_camera_settings(self):
        return self.config.get("camera", {})

    def get_excel_header(self):
        return self.config.get("excel_header", {})

    def get_log_path(self):
        return self.LOG_PATH

    def get_capture_dir(self):
        return self.CAPTURE_DIR
