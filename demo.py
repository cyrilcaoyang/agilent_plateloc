"""
demo.py — Agilent PlateLoc driver demonstration
=================================================

Run this after you have created a profile in the Diagnostics dialog
(with the correct COM port).  Change PROFILE below to match.

Usage:
    .venv\\Scripts\\python.exe demo.py
"""

import logging
import time

from agilent_plateloc import PlateLoc

# ── Configuration ───────────────────────────────────────────────────
PROFILE = "SDL2_PlateLoc"            # profile name you created in Diagnostics
COM_PORT = "COM14"             # serial port the PlateLoc is on
SEAL_TEMP_C = 170              # sealing temperature (°C)
SEAL_TIME_S = 1.2              # sealing duration  (seconds)
# ────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
)
log = logging.getLogger(__name__)


def main() -> None:
    sealer = PlateLoc(com_port=COM_PORT)

    # ── 1. Connect ──────────────────────────────────────────────────
    log.info("Connecting to PlateLoc (profile=%r) …", PROFILE)
    sealer.connect(profile=PROFILE)
    log.info("Connected!")

    # ── 2. Device info ──────────────────────────────────────────────
    print()
    print("═" * 50)
    print("  Agilent PlateLoc — Device Info")
    print("═" * 50)
    print(f"  ActiveX version  : {sealer.get_version()}")
    print(f"  Firmware version : {sealer.get_firmware_version()}")
    print(f"  Cycle count      : {sealer.get_cycle_count()}")
    print(f"  Profiles         : {sealer.enumerate_profiles()}")
    print("═" * 50)
    print()

    # ── 3. Read current settings ────────────────────────────────────
    cur_temp = sealer.get_sealing_temperature()
    cur_time = sealer.get_sealing_time()
    act_temp = sealer.get_actual_temperature()
    log.info("Current settings → temp=%d °C, time=%.1f s", cur_temp, cur_time)
    log.info("Actual plate temperature: %d °C", act_temp)

    # ── 4. Configure seal parameters ────────────────────────────────
    log.info("Setting temperature to %d °C …", SEAL_TEMP_C)
    sealer.set_sealing_temperature(SEAL_TEMP_C)

    log.info("Setting seal time to %.1f s …", SEAL_TIME_S)
    sealer.set_sealing_time(SEAL_TIME_S)

    # Confirm
    log.info(
        "Confirmed → temp=%d °C, time=%.1f s",
        sealer.get_sealing_temperature(),
        sealer.get_sealing_time(),
    )

    # ── 5. Wait for temperature to stabilise ────────────────────────
    log.info("Waiting for plate to reach %d °C …", SEAL_TEMP_C)
    for _ in range(120):  # up to ~2 minutes
        act = sealer.get_actual_temperature()
        if act >= SEAL_TEMP_C - 2:  # within 2 °C
            log.info("Plate ready at %d °C", act)
            break
        log.info("  heating … %d °C", act)
        time.sleep(1)
    else:
        log.warning("Timed out waiting for temperature — proceeding anyway")

    # ── 6. Run a seal cycle ─────────────────────────────────────────
    input("\n>>> Press ENTER to start a seal cycle (load plate first!) <<<\n")

    log.info("Starting seal cycle …")
    sealer.start_cycle()
    log.info("Seal cycle complete!")

    # ── 7. Read post-cycle info ─────────────────────────────────────
    log.info("Post-cycle temperature: %d °C", sealer.get_actual_temperature())
    log.info("Total cycle count: %d", sealer.get_cycle_count())
    last_err = sealer.get_last_error()
    if last_err:
        log.warning("Last error: %s", last_err)

    # ── 8. Disconnect ───────────────────────────────────────────────
    sealer.close()
    log.info("Done — PlateLoc disconnected.")


if __name__ == "__main__":
    main()
