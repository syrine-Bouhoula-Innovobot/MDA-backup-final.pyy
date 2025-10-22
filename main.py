   """
main.py
-------
Full main entry for the MDA Camera Automation project.
Includes the full run_sequence workflow (half-press → zoom → refocus → lock → capture groups).

Author: Syrine Bouhoula / MDA Project
Date: 2025
"""

import time
import edsdk
from edsdk import PropID, SaveTo

from config_manager import ConfigManager
from logger_excel import ExcelLogger
from error_handler import ErrorHandler
from edsdk_helpers import CameraController


def run_sequence(controller: CameraController, interactive=True):
    """
    Mirrors final.py logic:
      A) Half-press AF to wake
      B) Ramp to zoom 140
      C) Half-press AF again at 140
      D) Disable AF (lock focus)
    Then captures groups Z140, Z120, Z055, Z110 interactively.
    """

    print("📸 Initializing Canon SDK…")
    edsdk.InitializeSDK()

    try:
        # ---------- CAMERA CONNECTION ----------
        cl = edsdk.GetCameraList()
        if edsdk.GetChildCount(cl) == 0:
            print("❌ No camera detected.")
            return 1

        cam = edsdk.GetChildAtIndex(cl, 0)
        controller.cam_glob = cam

        print("✅ Camera connected. Opening session…")
        edsdk.OpenSession(cam)
        controller.pump(1.0)

        # --- UNLOCK AF STATE ---
        controller.reset_af_state(cam)
        controller.pump(0.6)

        # ---------- STORAGE SETTINGS ----------
        edsdk.SetPropertyData(cam, PropID.SaveTo, 0, SaveTo.Host)
        edsdk.SetCapacity(
            cam,
            {
                "reset": True,
                "bytesPerSector": 512,
                "numberOfFreeClusters": 2_147_483_647,
            },
        )
        print("💾 SaveTo=Host + capacity set.")

        # ---------- ENABLE EVF ----------
        controller.enable_evf(cam)

        # === A) Half-press AF to wake ===
        print("\n➡️ Step A: Half-press AF to wake…")
        if controller.half_press_with_retry(cam, controller.AF_WAIT_SEC):
            print("✅ Camera awake and focused (initial).")
        else:
            print("⚠️ Initial half-press failed.")
        controller.release_shutter(cam)
        controller.pump(0.4)

        # === B) Robust ramp to zoom 140 ===
        desc = controller.get_desc(cam, controller.pid_zoom)
        _, max_step = controller.derive_steps(desc)
        zoom_140 = min(int(controller.ZOOM_140_STR), max_step)
        print(f"\n🔍 Step B: Ramping to zoom {zoom_140} (no half-press held)…")
        ok, msg = controller.ramp_zoom(cam, controller.pid_zoom, zoom_140, quiet_settle=0.8, first=True)
        if ok:
            print("✅ Zoom ramp complete.")
        else:
            print("⚠️ Zoom ramp failed:", msg)
        print("🕐 Waiting 1.2 s for lens settle…")
        time.sleep(1.2)

        # === C) Half-press AF AGAIN at 140, then lock ===
        print("\n➡️ Step C: Half-press AF again at zoom 140…")
        if controller.half_press_with_retry(cam, wait=1.8):
            print("✅ Focus acquired at zoom 140.")
        else:
            print("⚠️ Focus attempt failed at zoom 140.")
        controller.release_shutter(cam)
        controller.pump(0.3)

        # === D) Disable all AF (manual-focus lock) ===
        print("\n🔒 Step D: Disabling all autofocus modes (locking focus)…")
        controller.disable_all_af_safe(cam)
        try:
            edsdk.SetPropertyData(cam, PropID.AFMode, 0, 3)  # Manual Focus (if supported)
            print("AFMode → 3 (Manual Focus).")
        except Exception:
            pass
        try:
            cur_ev = int(edsdk.GetPropertyData(cam, PropID.Evf_OutputDevice, 0))
            edsdk.SetPropertyData(cam, PropID.Evf_OutputDevice, 0, cur_ev)
            print("🔄 EVF refreshed to freeze focus state.")
        except Exception:
            pass
        print("✅ Focus locked — all subsequent shots will be NON-AF.\n")

        # ---------- ENABLE FILE TRANSFER CALLBACK ----------
        controller.enable_object_handler_after_ref(cam)

        # ---------- MANUAL CAPTURE SEQUENCE ----------
        print("\n--- Begin manual capture sequence ---")
        if interactive:
            print("Press ENTER when ready for each picture.\n")

        # ---- ZOOM 140 GROUP ----
        print("\n--- ZOOM 140 group ---")
        for cat in controller.ORDER_Z140:
            if interactive:
                input(f"\nReady for '{cat}' — press ENTER to capture…")
            controller.do_shot(
                cam,
                category=cat,
                tv_label_use=controller.TV_MAP.get(cat, controller.TV_REF_LABEL),
                is_reference=(cat == controller.CAT_REF),
                zoom_pid=controller.pid_zoom,
                zoom_target=None,
                av_label_use=controller.AV_LABEL,
                iso_label_use=controller.ISO_LABEL,
            )

        # ---- ZOOM 120 GROUP ----
        print("\n--- ZOOM 120 group ---")
        for idx, cat in enumerate(controller.ORDER_Z120):
            if interactive:
                input(f"\nReady for '{cat}' — press ENTER to capture…")
            controller.do_shot(
                cam,
                category=cat,
                tv_label_use=controller.TV_MAP.get(cat, controller.TV_REF_LABEL),
                is_reference=False,
                zoom_pid=controller.pid_zoom,
                zoom_target=120 if idx == 0 else None,
                av_label_use=controller.AV_LABEL,
                iso_label_use=controller.ISO_LABEL,
            )

        # ---- ZOOM 55 GROUP ----
        print("\n--- ZOOM 55 group ---")
        for idx, cat in enumerate(controller.ORDER_Z055):
            if interactive:
                input(f"\nReady for '{cat}' — press ENTER to capture…")
            controller.do_shot(
                cam,
                category=cat,
                tv_label_use=controller.TV_MAP.get(cat, controller.TV_REF_LABEL),
                is_reference=False,
                zoom_pid=controller.pid_zoom,
                zoom_target=55 if idx == 0 else None,
                av_label_use=controller.AV_LABEL,
                iso_label_use=controller.ISO_LABEL,
            )

        # ---- ZOOM 110 GROUP (LAST) ----
        print("\n--- ZOOM 110 group (LAST) ---")
        for cat in controller.ORDER_Z110:
            if interactive:
                input(f"\nReady for '{cat}' — press ENTER to capture…")
            controller.do_shot(
                cam,
                category=cat,
                tv_label_use=controller.TV_MAP.get(cat, controller.TV_REF_LABEL),
                is_reference=False,
                zoom_pid=controller.pid_zoom,
                zoom_target=110,
                av_label_use=controller.AV_LABEL,
                iso_label_use=controller.ISO_LABEL,
            )

        # ---------- FINALIZE ----------
        try:
            controller.reorder_logfile()
        except Exception as e:
            print("Reorder failed (non-fatal):", controller.safe_err_str(e))

        print("\n✅ All captures complete.")
        print("📂 Files saved under:", controller.capture_dir)
        print("📘 Log file used:", controller.log_path)

    finally:
        # ---------- SAFE EXIT ----------
        try:
            controller.release_all_locks(controller.cam_glob)
        except Exception:
            pass
        print("📴 Closing session…")
        controller.safe(edsdk.CloseSession, controller.cam_glob)
        print("🧩 Terminating SDK…")
        try:
            edsdk.TerminateSDK()
        finally:
            controller.pump(0.6)


def main():
    try:
        print("=== MDA Camera Automation — Starting session ===\n")

        cfg = ConfigManager("camera_config.json")
        logger = ExcelLogger(cfg.LOG_PATH, cfg.CAPTURE_DIR)
        controller = CameraController(cfg, logger)

        run_sequence(controller, interactive=True)

        print("\n✅ Sequence finished successfully.")

    except Exception as e:
        ErrorHandler.handle_error(e)
    finally:
        print("\n=== MDA Camera Automation — End of run ===")


if __name__ == "__main__":
    main()
