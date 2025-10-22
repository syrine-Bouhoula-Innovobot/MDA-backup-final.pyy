# edsdk_helpers.py
# ----------------
# Canon EDSDK helpers wrapped in a CameraController class (MDA project)
#
# Depends on:
#   - config_manager.ConfigManager   (you provide)
#   - logger_excel.ExcelLogger       (you provide)
#
# Author: MDA Project
# Date: 2025

import os
import time

import edsdk
from edsdk import (
    CameraCommand,
    PropID,
    SaveTo,
    ObjectEvent,
    FileCreateDisposition,
    Access,
    EdsObject,
)

# On Windows, we need to pump messages to keep the UI responsive
IS_WIN = (os.name == "nt")
if IS_WIN:
    import pythoncom


class CameraController:
    """
    Wraps all EDSDK-related helpers and the full capture pipeline
    that existed in final.py, but organized in a class.

    Usage:
        cam_ctl = CameraController(cfg, logger)
        cam_ctl.run_sequence()  # runs A-D focus/zoom sequence + capture groups
    """

    # ===== Tunables (same defaults as final.py) =====
    START_SETTLE_SEC = 1.2
    WRITE_RETRIES = 12
    POST_ZOOM_SETTLE = 0.7
    AF_WAIT_SEC = 1.1
    DEFAULT_DELTA = 3

    # Tv/Av lookup tables carried from final.py
    TV_LABELS = {
        24: '30"', 27: '25"', 29: '20"', 32: '15"', 35: '13"', 37: '10"', 40: '8"',
        43: '6"', 45: '5"', 48: '4"', 51: '3"2', 53: '2"5', 56: '2"', 59: '1"6',
        61: '1"3', 64: '1"', 67: '0"8', 69: '0"6', 72: '0"5', 75: '0"4', 77: '0"3',
        80: '1/5', 83: '1/6', 85: '1/8', 88: '1/10', 91: '1/20', 93: '1/25',
        96: '1/30', 99: '1/40', 101: '1/50', 104: '1/60', 107: '1/80', 109: '1/100',
        112: '1/125', 115: '1/160', 117: '1/200', 120: '1/250', 123: '1/320',
        125: '1/400', 128: '1/500', 131: '1/640', 133: '1/800', 136: '1/1000',
        139: '1/1250', 141: '1/1600', 144: '1/2000'
    }
    AV_LABELS = {
        40: "f/3.5", 43: "f/4.0", 45: "f/4.5", 48: "f/5.0", 51: "f/5.6",
        53: "f/7.1", 56: "f/8", 59: "f/9.0", 61: "f/10", 64: "f/11", 67: "f/13"
    }

    # Shutter press codes
    SH_OFF = 0x00000000
    SH_HALF = 0x00000001
    SH_FULL_NONAF = 0x00010003  # full press without AF

    def __init__(self, cfg, logger):
        """
        cfg: ConfigManager instance
        logger: ExcelLogger instance
        """
        self.cfg = cfg
        self.logger = logger

        # Working/session info
        self.project_id = getattr(cfg, "PROJECT_ID", "MDA12345")
        self.device_code = getattr(cfg, "DEVICE_CODE", "DC123")
        self.capture_dir = getattr(cfg, "CAPTURE_DIR", os.getcwd())
        self.log_path = getattr(cfg, "LOG_PATH", os.path.join(self.capture_dir, "log.xlsx"))

        # Camera parameters from config
        self.AV_LABEL = getattr(cfg, "AV_LABEL", "f/8")
        self.ISO_LABEL = getattr(cfg, "ISO_LABEL", "100")
        self.TV_REF_LABEL = getattr(cfg, "TV_REF_LABEL", "1/60")
        self.POST_SHOT_WAIT = getattr(cfg, "POST_SHOT_WAIT", 2.5)
        self.THUMBNAIL_WIDTH_PX = getattr(cfg, "THUMBNAIL_WIDTH_PX", 180)

        # Zoom targets as strings/numbers from config
        self.ZOOM_140_STR = getattr(cfg, "ZOOM_140_STR", "140")
        self.ZOOM_120_STR = getattr(cfg, "ZOOM_120_STR", "120")
        self.ZOOM_110_STR = getattr(cfg, "ZOOM_110_STR", "110")
        self.ZOOM_055_STR = getattr(cfg, "ZOOM_055_STR", "55")

        # Category + order setup
        self.CAT_REF = getattr(cfg, "CAT_REF", "reference focus sticker")
        self.FEATURE_ORDER = getattr(cfg, "FEATURE_ORDER", [])
        self.ORDER_Z140 = getattr(cfg, "ORDER_Z140", [])
        self.ORDER_Z120 = getattr(cfg, "ORDER_Z120", [])
        self.ORDER_Z055 = getattr(cfg, "ORDER_Z055", [])
        self.ORDER_Z110 = getattr(cfg, "ORDER_Z110", [])

        # Per-feature Tv mapping
        self.TV_MAP = getattr(cfg, "TV_MAP", {})

        # Runtime state
        self.cam_glob = None
        self.CURRENT_CATEGORY = self.CAT_REF
        self.saved_paths = []
        self.OBJ_HANDLER_ENABLED = False

        # Resolve zoom property id once
        self.pid_zoom = self.dc_zoom_id()

    # ------------------- small utilities -------------------

    @staticmethod
    def safe_err_str(e):
        try:
            return str(e)
        except Exception:
            try:
                if hasattr(e, "args") and e.args:
                    return str(e.args[0])
            except Exception:
                pass
            try:
                return repr(e)
            except Exception:
                return "<unprintable error>"

    @staticmethod
    def err_is_busy(e):
        s = CameraController.safe_err_str(e).lower()
        return ("device_busy" in s) or ("eds_err_device_busy" in s) or ("busy" in s)

    def set_category(self, cat: str):
        self.CURRENT_CATEGORY = str(cat)

    def pump(self, total=0.35, dt=0.05):
        t0 = time.time()
        while time.time() - t0 < total:
            try:
                edsdk.SendCommand(self.cam_glob, CameraCommand.ExtendShutDownTimer, 0)
            except Exception:
                pass
            if IS_WIN:
                try:
                    pythoncom.PumpWaitingMessages()
                except Exception:
                    pass
            time.sleep(dt)

    @staticmethod
    def safe(cmd, *args):
        try:
            cmd(*args)
        except Exception:
            pass

    @staticmethod
    def dc_zoom_id():
        try:
            return getattr(edsdk, "kEdsPropID_DC_Zoom")
        except Exception:
            return getattr(PropID, "DC_Zoom")

    @staticmethod
    def status_enum():
        try:
            from edsdk import StatusCommand
            return StatusCommand
        except Exception:
            return None

    # ------------------- EVF + keepalive -------------------

    def keepalive_no_half(self, cam):
        self.safe(edsdk.SendCommand, cam, CameraCommand.ExtendShutDownTimer, 0)
        self.pump(0.15)
        try:
            cur = int(edsdk.GetPropertyData(cam, PropID.Evf_OutputDevice, 0))
            self.safe(edsdk.SetPropertyData, cam, PropID.Evf_OutputDevice, 0, cur)
        except Exception:
            pass
        self.pump(0.10)
        for pid in (PropID.Tv, PropID.Av, PropID.ISOSpeed):
            try:
                _ = edsdk.GetPropertyData(cam, pid, 0)
            except Exception:
                pass
        self.pump(0.10)

    def enable_evf(self, cam):
        self.safe(edsdk.SetPropertyData, cam, PropID.Evf_Mode, 0, 1)
        self.pump(0.18)
        self.safe(edsdk.SetPropertyData, cam, PropID.Evf_OutputDevice, 0, 2)
        self.pump(0.22)
        print("EVF mode -> ON\nEVF output -> PC")

    # ------------------- Zoom helpers -------------------

    @staticmethod
    def get_desc(cam, pid):
        try:
            return edsdk.GetPropertyDesc(cam, pid) or {}
        except Exception as e:
            print("GetPropertyDesc failed:", e)
            return {}

    @staticmethod
    def derive_steps(desc):
        total, max_step = None, None
        if isinstance(desc, dict):
            pd = desc.get("propDesc")
            if isinstance(pd, (list, tuple)) and pd:
                if len(pd) == 1 and isinstance(pd[0], int):
                    total = int(pd[0])
                    max_step = total - 1
                else:
                    vals = [int(x) for x in pd if isinstance(x, int)]
                    if vals:
                        max_step = max(vals)
                        total = max_step - min(vals) + 1
            if total is None:
                for k in ("max", "Max", "maximum", "Maximum"):
                    if k in desc and isinstance(desc[k], int):
                        max_step = int(desc[k])
                        total = max_step + 1
                        break
        return total or 201, max_step or 200

    @staticmethod
    def read_zoom(cam, pid):
        try:
            return int(edsdk.GetPropertyData(cam, pid, 0))
        except Exception:
            return None

    def write_zoom(self, cam, pid, step, StatusCommand=None, retries=None):
        if retries is None:
            retries = self.WRITE_RETRIES
        step = int(step)
        last = ""
        for i in range(retries):
            try:
                edsdk.SetPropertyData(cam, pid, 0, step)
                self.pump(0.24 + 0.12 * i)
                rb = self.read_zoom(cam, pid)
                if rb == step:
                    return True
            except Exception as e:
                last = self.safe_err_str(e)
            self.pump(0.28 + 0.14 * i)
        if last:
            print("Set DC_Zoom failed:", last)
        return False

    def reset_liveview_for_zoom(self, cam):
        """Bootstrap LV so the FIRST ramp accepts writes."""
        try:
            self.safe(edsdk.SendCommand, cam, CameraCommand.PressShutterButton, 0)
            self.keepalive_no_half(cam)
            self.pump(0.20)
            try:
                self.safe(edsdk.SetPropertyData, cam, PropID.Evf_Mode, 0, 1)
            except Exception:
                pass
            self.pump(0.18)
            self.safe(edsdk.SetPropertyData, cam, PropID.Evf_OutputDevice, 0, 0)  # off
            self.pump(0.45)
            self.safe(edsdk.SetPropertyData, cam, PropID.Evf_OutputDevice, 0, 2)  # PC
            self.pump(1.0)
            for pid in (PropID.Tv, PropID.Av, PropID.ISOSpeed):
                try:
                    _ = edsdk.GetPropertyData(cam, pid, 0)
                except Exception:
                    pass
            self.pump(0.30)
        except Exception as e:
            print("reset_liveview_for_zoom: ignored error:", self.safe_err_str(e))

    def ramp_zoom(self, cam, pid, target, delta=None, quiet_settle=0.0, first=False):
        """
        Robust zoom ramp that NEVER holds half-press.
        UI lock is held for the entire ramp to reduce BUSY flicker.
        If it fails at the beginning, reset LV once and retry.
        """
        if delta is None:
            delta = self.DEFAULT_DELTA

        se = self.status_enum()

        def _do_ramp():
            # Make sure nothing is half-pressed.
            self.safe(edsdk.SendCommand, cam, CameraCommand.PressShutterButton, 0)
            self.keepalive_no_half(cam)
            self.pump(0.25)

            cur = self.read_zoom(cam, pid)
            if cur is None:
                print("ramp_zoom: cannot read zoom.")
                return False, "cannot read zoom"
            if target == cur:
                print(f"ramp_zoom: already at {cur}")
                self.pump(self.POST_ZOOM_SETTLE)
                return True, cur

            if se:
                self.safe(edsdk.SendStatusCommand, cam, se.UILock)
            try:
                sgn = 1 if target > cur else -1
                step = cur
                while step != target:
                    step = step + sgn * min(delta, abs(target - step))
                    ok = self.write_zoom(cam, pid, step, StatusCommand=None)
                    print(f"  zoom write -> {step} : {'OK' if ok else 'FAIL'}")
                    if not ok:
                        return False, f"write failed at step {step}"
                    self.pump(0.08)
            finally:
                if se:
                    self.safe(edsdk.SendStatusCommand, cam, se.UIUnLock)
            self.pump(self.POST_ZOOM_SETTLE)
            return True, step

        ok, msg = _do_ramp()
        if not ok:
            print("ramp_zoom: first attempt failed; resetting Live View and retrying…")
            self.reset_liveview_for_zoom(cam)
            self.pump(0.35)
            ok, msg = _do_ramp()

        if ok and quiet_settle > 0:
            print(f"Quiet settle {quiet_settle:.1f}s…")
            self.pump(quiet_settle)

        return ok, msg

    # ------------------- Shutter + AF helpers -------------------

    def release_shutter(self, cam):
        self.safe(edsdk.SendCommand, cam, CameraCommand.PressShutterButton, self.SH_OFF)

    def half_press_with_retry(self, cam, wait=None, retries=6):
        """Used for AF before we disable it. Busy-safe with backoff."""
        if wait is None:
            wait = self.AF_WAIT_SEC
        self.keepalive_no_half(cam)  # wake without half-press
        self.pump(0.25)
        for i in range(retries):
            try:
                edsdk.SendCommand(cam, CameraCommand.PressShutterButton, self.SH_HALF)
                self.pump(wait)
                return True
            except Exception as e:
                if self.err_is_busy(e):
                    back = 0.35 + 0.25 * i
                    print(f"Half-press busy, retrying in {back:.2f}s…")
                    self.pump(back)
                    self.keepalive_no_half(cam)
                    continue
                print("Half-press failed:", self.safe_err_str(e))
                return False
        print("Half-press failed: device busy (exhausted retries).")
        return False

    def capture_full_nonaf(self, cam, hold=0.5, retries=6):
        for i in range(retries):
            try:
                edsdk.SendCommand(cam, CameraCommand.PressShutterButton, self.SH_FULL_NONAF)
                self.pump(hold)
                self.release_shutter(cam)
                return True
            except Exception as e:
                if self.err_is_busy(e):
                    self.pump(0.25 + 0.25 * i)
                    self.safe(edsdk.SendCommand, cam, CameraCommand.ExtendShutDownTimer, 0)
                    continue
                print("NON-AF capture failed:", self.safe_err_str(e))
                self.release_shutter(cam)
                return False
        print("NON-AF capture failed: device busy.")
        self.release_shutter(cam)
        return False

    def disable_all_af_safe(self, cam):
        """Disable all AF systems (PowerShot-safe)."""
        print("\n🔒 Disabling autofocus (PowerShot-safe)…")
        try:
            from edsdk import StatusCommand
            se = StatusCommand
        except Exception:
            se = None

        def set_prop(pid_name, value, label=""):
            pid = getattr(PropID, pid_name, None)
            if pid is None:
                return
            if se:
                self.safe(edsdk.SendStatusCommand, cam, se.UILock)
            try:
                edsdk.SetPropertyData(cam, pid, 0, value)
                if label:
                    print(f"{label} -> {value}")
            except Exception as e:
                if label:
                    print(f"{label} set failed:", self.safe_err_str(e))
            finally:
                if se:
                    self.safe(edsdk.SendStatusCommand, cam, se.UIUnLock)
            self.pump(0.15)

        props = [
            ("ContinuousAFMode", 0, "ContinuousAFMode=Off"),
            ("MovieServoAF", 0, "MovieServoAF=Off"),
            ("LensDriveWhenAFImpossible", 0, "LensDriveWhenAFImpossible=Off"),
            ("AFAssist", 0, "AFAssist=Off"),
        ]
        for n, v, l in props:
            set_prop(n, v, l)

        # Refresh EVF to apply “pseudo-MF” freeze
        try:
            cur = int(edsdk.GetPropertyData(cam, PropID.Evf_OutputDevice, 0))
            edsdk.SetPropertyData(cam, PropID.Evf_OutputDevice, 0, cur)
        except Exception:
            pass
        print("✅ AF systems off. Focus should remain fixed.\n")

    def reset_af_state(self, cam):
        """Re-enable all autofocus systems and clear any half-press lock before the main sequence."""
        print("\n🔄 Resetting AF / Focus system (unlock half-press)…")

        # Enable AF-related properties
        for pid_name in ("ContinuousAFMode", "MovieServoAF", "LensDriveWhenAFImpossible", "AFAssist"):
            pid = getattr(PropID, pid_name, None)
            if pid is None:
                continue
            try:
                edsdk.SetPropertyData(cam, pid, 0, 1)
                print(f"{pid_name} → 1 (enabled)")
            except Exception as e:
                print(f"Could not enable {pid_name}:", self.safe_err_str(e))

        # Set AF mode to One-Shot (0) if supported
        try:
            edsdk.SetPropertyData(cam, PropID.AFMode, 0, 0)
            print("AFMode → 0 (One-Shot AF)")
        except Exception as e:
            print("Could not set AFMode:", self.safe_err_str(e))

        # Refresh EVF and run quick AF pulse
        try:
            edsdk.SetPropertyData(cam, PropID.Evf_Mode, 0, 1)
            edsdk.SetPropertyData(cam, PropID.Evf_OutputDevice, 0, 2)
            edsdk.SendCommand(cam, CameraCommand.DoEvfAf, 1)
            self.pump(0.5)
            edsdk.SendCommand(cam, CameraCommand.DoEvfAf, 0)
            print("📸 AF pulse sent to re-initialize lens drive.")
        except Exception:
            pass

        print("✅ AF state unlocked — half-press should now work normally.\n")

    def release_all_locks(self, cam):
        self.release_shutter(cam)

    # ------------------- Save-to-PC + callbacks -------------------

    @staticmethod
    def ensure_dir(path):
        os.makedirs(path, exist_ok=True)

    def category_dir(self):
        return self.capture_dir

    def _save_item(self, handle: EdsObject, out_dir: str):
        info = edsdk.GetDirectoryItemInfo(handle)
        try:
            cat_prefix = f"{self.project_id}_{self.CURRENT_CATEGORY.replace(' ', '_')}"
            base_name = f"{cat_prefix}_{time.strftime('%H%M%S')}"
            out_path = os.path.join(out_dir, base_name + ".JPG")

            stream = edsdk.CreateFileStream(out_path, FileCreateDisposition.CreateAlways, Access.ReadWrite)
            edsdk.Download(handle, info["size"], stream)
            edsdk.DownloadComplete(handle)
            self.saved_paths.append(out_path)
            print("Saved:", out_path)
        except Exception as e:
            print("Save failed:", self.safe_err_str(e))

    # NOTE: EDSDK expects a plain function. We provide a bound dispatcher via static wrapper.
    def _obj_callback_impl(self, event: ObjectEvent, handle: EdsObject) -> int:
        try:
            if event in (ObjectEvent.DirItemRequestTransfer, ObjectEvent.DirItemCreated):
                self._save_item(handle, self.category_dir())
        except Exception as e:
            print("Save (callback) failed:", self.safe_err_str(e))
        return 0

    def enable_object_handler_after_ref(self, cam):
        if self.OBJ_HANDLER_ENABLED:
            return
        try:
            # We create a closure to bind "self" to a static function compatible with EDSDK.
            def _cb(event, handle):
                return self._obj_callback_impl(event, handle)
            edsdk.SetObjectEventHandler(cam, ObjectEvent.All, _cb)
            self.OBJ_HANDLER_ENABLED = True
            print("ObjectEvent handler enabled AFTER reference shot.")
        except Exception as e:
            print("ObjectEvent handler enable failed:", self.safe_err_str(e))

    @staticmethod
    def list_jpgs(path):
        if not os.path.isdir(path):
            return []
        return [
            os.path.join(path, f)
            for f in os.listdir(path)
            if f.lower().endswith((".jpg", ".jpeg"))
        ]

    # ------------------- Tv / Av / ISO / EC helpers -------------------

    @classmethod
    def tv_label(cls, code):
        return cls.TV_LABELS.get(int(code), f"code={int(code)}")

    @staticmethod
    def tv_desc(cam):
        try:
            d = edsdk.GetPropertyDesc(cam, PropID.Tv) or {}
            return [int(v) for v in d.get("propDesc") or []]
        except Exception:
            return []

    @classmethod
    def parse_label_to_index(cls, vals, want_label):
        want = want_label.strip()
        for i, c in enumerate(vals):
            if cls.tv_label(c) == want or str(c) == want:
                return i
        raise ValueError("Tv label not in allowed list")

    @staticmethod
    def read_tv(cam):
        try:
            return int(edsdk.GetPropertyData(cam, PropID.Tv, 0))
        except Exception:
            return None

    def set_tv_by_index(self, cam, idx, vals):
        code = int(vals[idx])
        edsdk.SetPropertyData(cam, PropID.Tv, 0, code)
        self.pump(0.6)
        rb = self.read_tv(cam)
        print(f"Tv: set code={code} ({self.tv_label(code)}) -> readback {rb} ({self.tv_label(rb)})")
        return code, rb

    def set_tv_preferred(self, cam, want_label):
        vals = self.tv_desc(cam)
        if not vals:
            print("No Tv list; leaving Tv as-is.")
            return None, self.read_tv(cam)
        lookup = {"1/25": 93, "1/20": 91, "1/100": 109, "1/30": 96, "1/50": 101,
                  "1/80": 107, "1/40": 99, "1/125": 112, "1/320": 123}
        try:
            idx = self.parse_label_to_index(vals, want_label)
            return self.set_tv_by_index(cam, idx, vals)
        except Exception:
            target_code = lookup.get(want_label, 93)
            idx = min(range(len(vals)), key=lambda i: abs(int(vals[i]) - target_code))
            return self.set_tv_by_index(cam, idx, vals)

    @classmethod
    def av_label(cls, code):
        return cls.AV_LABELS.get(int(code), f"code={int(code)}")

    @staticmethod
    def av_desc(cam):
        try:
            d = edsdk.GetPropertyDesc(cam, PropID.Av) or {}
            return [int(v) for v in d.get("propDesc") or []]
        except Exception:
            return []

    @classmethod
    def choose_av_code(cls, vals, want_label):
        want = (want_label or "").replace(".0", "")
        for c in vals:
            if cls.av_label(c).replace(".0", "") == want:
                return int(c)
        return int(min(vals, key=lambda c: abs(int(c) - 56))) if vals else None

    def set_av(self, cam, code):
        edsdk.SetPropertyData(cam, PropID.Av, 0, int(code))
        self.pump(0.5)
        rb = int(edsdk.GetPropertyData(cam, PropID.Av, 0))
        print(f"Av: set code={code} ({self.av_label(code)}) -> readback {rb} ({self.av_label(rb)})")
        return rb

    @staticmethod
    def _iso_desc(cam):
        try:
            d = edsdk.GetPropertyDesc(cam, PropID.ISOSpeed) or {}
            return [int(v) for v in d.get("propDesc") or []]
        except Exception:
            return []

    _STD_ISO_SERIES = [100, 125, 160, 200, 250, 320, 400, 500, 640, 800, 1000, 1250, 1600, 2000, 2500, 3200]

    @classmethod
    def _iso_series_for_desc(cls, desc):
        nz = [c for c in desc if c != 0]
        return dict(zip(nz, cls._STD_ISO_SERIES[:len(nz)]))

    def iso_label_from_code(self, cam, code):
        desc = self._iso_desc(cam)
        if not desc:
            return "Auto" if code == 0 else f"code={int(code)}"
        if code == 0:
            return "Auto"
        return str(self._iso_series_for_desc(desc).get(int(code), f"code={int(code)}"))

    def set_iso_preferred(self, cam, want_label: str):
        desc = self._iso_desc(cam)
        if not desc:
            print("ISO list not exposed; skipping ISO set.")
            try:
                return int(edsdk.GetPropertyData(cam, PropID.ISOSpeed, 0))
            except Exception:
                return None
        map_by_order = self._iso_series_for_desc(desc)
        inv = {v: k for k, v in map_by_order.items()}
        s = str(want_label).strip().lower().replace("iso", "")
        if s in ("auto", "0"):
            target_code = 0 if 0 in desc else None
        else:
            try:
                want_num = int(s)
            except Exception:
                want_num = None
            if want_num in inv:
                target_code = inv[want_num]
            else:
                if want_num is None or not inv:
                    target_code = None
                else:
                    nearest = min(map_by_order.values(), key=lambda x: abs(x - want_num))
                    target_code = inv[nearest]
        if target_code is None:
            print("ISO: could not resolve requested value; leaving as-is.")
            try:
                return int(edsdk.GetPropertyData(cam, PropID.ISOSpeed, 0))
            except Exception:
                return None
        edsdk.SetPropertyData(cam, PropID.ISOSpeed, 0, int(target_code))
        self.pump(0.5)
        rb = int(edsdk.GetPropertyData(cam, PropID.ISOSpeed, 0))
        print(
            f"ISO: set code={target_code} ({self.iso_label_from_code(cam, target_code)}) -> "
            f"readback {rb} ({self.iso_label_from_code(cam, rb)})"
        )
        return rb

    def read_iso_label(self, cam):
        try:
            code = int(edsdk.GetPropertyData(cam, PropID.ISOSpeed, 0))
            return self.iso_label_from_code(cam, code)
        except Exception:
            return ""

    @staticmethod
    def _ec_desc_list(cam):
        try:
            d = edsdk.GetPropertyDesc(cam, PropID.ExposureCompensation) or {}
            return [int(v) for v in d.get("propDesc") or []]
        except Exception:
            return []

    def read_ec(self, cam):
        try:
            cur_code = int(edsdk.GetPropertyData(cam, PropID.ExposureCompensation, 0))
        except Exception:
            return ""
        vals = self._ec_desc_list(cam)
        if not vals or cur_code not in vals:
            return 0.0 if cur_code == 0 else ""
        z = vals.index(0)
        i = vals.index(cur_code)
        steps = i - z
        whole, rem = divmod(abs(steps), 3)
        frac = {0: 0.0, 1: 0.3, 2: 0.7}[rem]
        v = (whole + frac) * (1 if steps >= 0 else -1)
        return float(f"{max(-3.0, min(3.0, v)):.1f}")


    # ------------------- Shot pipeline -------------------

    def do_shot(self, cam, *, category, tv_label_use, is_reference, zoom_pid, zoom_target, av_label_use, iso_label_use):
        # Update current category for saved filename prefix
        self.set_category(category)

        folder_before = set(self.list_jpgs(self.category_dir()))

        # Tv first (quiet)
        _res = self.set_tv_preferred(cam, tv_label_use)
        tv_code = _res[1] if isinstance(_res, tuple) else self.read_tv(cam)
        tv_str = self.tv_label(tv_code) if tv_code is not None else ""

        # Zoom (if requested)
        if zoom_target is not None:
            desc = self.get_desc(cam, zoom_pid)
            total, max_step = self.derive_steps(desc)
            cur = self.read_zoom(cam, zoom_pid)
            print(f"DC_Zoom: total={total}, range 0..{max_step}, current={cur}, target={zoom_target}")
            ok, msg = self.ramp_zoom(cam, zoom_pid, zoom_target, quiet_settle=0.0, first=False)
            if not ok:
                print("Zoom ramp FAILED:", msg)
            rbz = self.read_zoom(cam, zoom_pid)
            print("Zoom readback:", rbz)
        else:
            rbz = self.read_zoom(cam, zoom_pid)
            print("Zoom unchanged at:", rbz)

        # Av / ISO
        av_vals = self.av_desc(cam)
        if av_vals:
            av_code = self.choose_av_code(av_vals, av_label_use)
            if av_code is not None:
                self.set_av(cam, av_code)
        try:
            av_code_now = int(edsdk.GetPropertyData(cam, PropID.Av, 0))
        except Exception:
            av_code_now = None
        av_str = self.av_label(av_code_now) if av_code_now is not None else ""

        iso_code_now = self.set_iso_preferred(cam, iso_label_use)
        iso_str = self.iso_label_from_code(cam, iso_code_now) if iso_code_now is not None else self.read_iso_label(cam)

        # Shoot (always NON-AF after we lock)
        start_idx = len(self.saved_paths)
        self.pump(0.25)
        if is_reference:
            print("Reference capture (NON-AF)…")
            self.capture_full_nonaf(cam)
        else:
            print("Capture (NON-AF)…")
            self.capture_full_nonaf(cam)

        # Wait + gather new files
        self.pump(self.POST_SHOT_WAIT)
        new_files = self.saved_paths[start_idx:]

        # Fallback scan
        if not new_files:
            folder_after = set(self.list_jpgs(self.category_dir()))
            diff = sorted(folder_after - folder_before, key=os.path.getmtime)
            if diff:
                print(f"Fallback picked up {len(diff)} file(s).")
                new_files = diff
                self.saved_paths.extend([p for p in diff if p not in self.saved_paths])

        ec_val = self.read_ec(cam)

        if new_files:
            for idx, path in enumerate(new_files, start=1):
                # In final.py, Excel rows: [ts, cat, picture, path, tv, av, zoom, iso, ec]
                ts = time.strftime("%Y-%m-%d %H:%M:%S")
                # We keep the logger interface you already created:
                # logger.append_row(order_idx, category, image_path, tv_str, av_str, zoom, iso_str, ec_val)
                self.logger.append_row(
                    order_idx=None,  # if you want to fill order externally; left None here
                    category=category,
                    image_path=path,
                    tv_str=self.tv_text_for_log(tv_str),
                    av_str=av_str,
                    zoom=str(rbz if rbz is not None else ""),
                    iso_str=iso_str,
                    ec_val=ec_val,
                )
            print(f"Logged {len(new_files)} shot(s) for '{category}'.")
        else:
            print("No files detected to log for this shot.")

    @staticmethod
    def tv_text_for_log(tv_str):
        return f"{tv_str} s" if tv_str and "/" in tv_str and '"' not in tv_str else tv_str

 