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
from agilent_plateloc.plateloc import PlateLocError
from agilent_plateloc.config import (
    get as cfg,
    get_seal_params,
    load_film_settings,
)

# ── Configuration (loaded from config.toml) ────────────────────────
PROFILE   = cfg("instrument", "profile",  "default")
COM_PORT  = cfg("instrument", "com_port", "COM4")
DEF_SEAL_NAME = cfg("film", "seal_name", "Peelable Aluminum")
DEF_PLATE_MATERIAL = cfg("film", "plate_material", "polypropylene")
TEMP_TOL_C  = int(cfg("film", "temperature_tolerance_c", 2))
HEAT_TIMEOUT = int(cfg("film", "heat_timeout_s", 120))
# ────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
)
log = logging.getLogger(__name__)


def _choose_film_and_material() -> tuple[str, str]:
    """Let the user pick a seal film and plate material from film_settings.json."""
    data = load_film_settings()
    films = data.get("seal_films") or []
    if not films:
        raise SystemExit("No seal_films defined in film_settings.json")

    print()
    print("=" * 50)
    print("Available seal films:")
    for idx, film in enumerate(films, start=1):
        name = film.get("name", "<unnamed>")
        pn = film.get("product_number", "")
        label = f"{name}"
        if pn:
            label += f" (PN {pn})"
        print(f"  [{idx}] {label}")
    print("=" * 50)

    # Film selection (by index, default from config)
    try:
        default_idx = next(
            i for i, f in enumerate(films, start=1)
            if f.get("name") == DEF_SEAL_NAME or f.get("product_number") == DEF_SEAL_NAME
        )
    except StopIteration:
        default_idx = 1

    raw = input(f"Select seal film [1-{len(films)}] (default {default_idx}): ").strip()
    if not raw:
        film_idx = default_idx
    else:
        film_idx = int(raw)
        if film_idx < 1 or film_idx > len(films):
            raise SystemExit("Invalid film selection")

    film = films[film_idx - 1]
    film_name = film.get("name") or film.get("product_number")

    # Plate material selection
    compat = film.get("microplate_compatibility") or {}
    if not isinstance(compat, dict) or not compat:
        raise SystemExit(f"No microplate_compatibility defined for film {film_name!r}")

    materials: list[str] = []
    for mat, info in compat.items():
        if isinstance(info, dict):
            materials.append(mat)
    if not materials:
        raise SystemExit(f"No compatible plate materials for film {film_name!r}")

    print()
    print("Compatible plate materials for this film:")
    for idx, mat in enumerate(materials, start=1):
        print(f"  [{idx}] {mat}")

    try:
        default_mat_idx = materials.index(DEF_PLATE_MATERIAL) + 1
    except ValueError:
        default_mat_idx = 1

    raw = input(
        f"Select plate material [1-{len(materials)}] (default {default_mat_idx}): "
    ).strip()
    if not raw:
        mat_idx = default_mat_idx
    else:
        mat_idx = int(raw)
        if mat_idx < 1 or mat_idx > len(materials):
            raise SystemExit("Invalid plate material selection")

    plate_material = materials[mat_idx - 1]
    return film_name, plate_material


def _prompt_int_setting(label: str, default: int, min_value: int, max_value: int) -> int:
    """Prompt for an integer setting, using default on blank input."""
    while True:
        raw = input(
            f"{label} [{min_value}-{max_value}] (default {default}): "
        ).strip()
        if not raw:
            return default

        try:
            value = int(raw)
        except ValueError:
            print("Please enter a whole number.")
            continue

        if min_value <= value <= max_value:
            return value
        print(f"Please enter a value from {min_value} to {max_value}.")


def _prompt_float_setting(
    label: str,
    default: float,
    min_value: float,
    max_value: float,
) -> float:
    """Prompt for a floating-point setting, using default on blank input."""
    while True:
        raw = input(
            f"{label} [{min_value:.1f}-{max_value:.1f}] "
            f"(default {default:.1f}): "
        ).strip()
        if not raw:
            return default

        try:
            value = float(raw)
        except ValueError:
            print("Please enter a number.")
            continue

        if min_value <= value <= max_value:
            return value
        print(f"Please enter a value from {min_value:.1f} to {max_value:.1f}.")


def _customize_seal_params(default_temp_c: int, default_time_s: float) -> tuple[int, float]:
    """Allow operator override of film-derived sealing parameters."""
    print()
    print("=" * 50)
    print("Seal parameters")
    print("=" * 50)
    print(f"Recommended temperature : {default_temp_c} C")
    print(f"Recommended seal time   : {default_time_s:.1f} s")
    print("Press ENTER to use each recommended value, or type a custom value.")
    print()

    seal_temp_c = _prompt_int_setting("Sealing temperature (C)", default_temp_c, 20, 235)
    seal_time_s = _prompt_float_setting("Sealing time (s)", default_time_s, 0.5, 12.0)
    return seal_temp_c, seal_time_s


def _wait_for_temperature_ready(sealer: PlateLoc, target_c: int) -> None:
    """Block until the plate is within tolerance, or abort before cycling."""
    log.info(
        "Waiting for plate to reach %d C (+/- %d C). Press Ctrl+C to abort ...",
        target_c,
        TEMP_TOL_C,
    )
    try:
        for _ in range(HEAT_TIMEOUT):
            act = sealer.get_actual_temperature()
            delta = act - target_c
            if abs(delta) <= TEMP_TOL_C:
                log.info("Plate ready at %d C", act)
                return

            direction = "cooling" if delta > 0 else "heating"
            log.info("  %s ... %d C", direction, act)
            time.sleep(1)
    except KeyboardInterrupt:
        log.warning("Temperature wait aborted by operator; seal cycle will not start")
        sealer.close()
        raise SystemExit(1) from None

    actual_c = sealer.get_actual_temperature()
    log.error(
        "Timed out waiting for plate temperature; seal cycle will not start "
        "(actual=%d C, target=%d C, tolerance=%d C)",
        actual_c,
        target_c,
        TEMP_TOL_C,
    )
    sealer.close()
    raise SystemExit(1)


def main() -> None:
    sealer = PlateLoc(com_port=COM_PORT)

    # ── 1. Connect ──────────────────────────────────────────────────
    log.info("Connecting to PlateLoc (profile=%r) ...", PROFILE)
    sealer.connect(profile=PROFILE)
    log.info("Connected!")

    # ── 2. Device info ──────────────────────────────────────────────
    print()
    print("=" * 50)
    print("  Agilent PlateLoc -- Device Info")
    print("=" * 50)
    print(f"  ActiveX version  : {sealer.get_version()}")
    print(f"  Firmware version : {sealer.get_firmware_version()}")
    print(f"  Cycle count      : {sealer.get_cycle_count()}")
    print(f"  Profiles         : {sealer.enumerate_profiles()}")
    print("=" * 50)
    print()

    # ── 3. Let user choose seal film / plate material ──────────────
    film_name, plate_material = _choose_film_and_material()
    params = get_seal_params(film_name, plate_material)
    recommended_temp_c = int(params["temperature_c"])
    recommended_time_s = float(params["time_s"])
    seal_temp_c, seal_time_s = _customize_seal_params(
        recommended_temp_c,
        recommended_time_s,
    )

    log.info(
        "Using film=%r, plate=%r -> temp=%d C, time=%.2f s",
        film_name,
        plate_material,
        seal_temp_c,
        seal_time_s,
    )

    # ── 4. Read current settings ────────────────────────────────────
    cur_temp = sealer.get_sealing_temperature()
    cur_time = sealer.get_sealing_time()
    act_temp = sealer.get_actual_temperature()
    log.info("Current settings -> temp=%d C, time=%.1f s", cur_temp, cur_time)
    log.info("Actual plate temperature: %d C", act_temp)

    # ── 5. Configure seal parameters ────────────────────────────────
    log.info("Setting temperature to %d C ...", seal_temp_c)
    sealer.set_sealing_temperature(seal_temp_c)

    log.info("Setting seal time to %.1f s ...", seal_time_s)
    sealer.set_sealing_time(seal_time_s)

    # Confirm
    log.info(
        "Confirmed -> temp=%d C, time=%.1f s",
        sealer.get_sealing_temperature(),
        sealer.get_sealing_time(),
    )

    # ── 6. Wait for temperature to stabilise ────────────────────────
    _wait_for_temperature_ready(sealer, seal_temp_c)

    # ── 7. Run a seal cycle ─────────────────────────────────────────
    input("\n>>> Press ENTER to start a seal cycle (load plate first!) <<<\n")

    log.info("Starting seal cycle ...")
    try:
        sealer.start_cycle()
        log.info("Seal cycle complete!")
    except PlateLocError:
        err_msg = sealer.get_last_error()
        log.error("Seal cycle FAILED: %s", err_msg)
        log.info("Aborting current operation ...")
        try:
            sealer.abort()
            log.info("Abort sent successfully")
        except PlateLocError:
            pass  # abort may also fail if already cleared

    # ── 8. Read post-cycle info ─────────────────────────────────────
    log.info("Post-cycle temperature: %d C", sealer.get_actual_temperature())
    log.info("Total cycle count: %d", sealer.get_cycle_count())
    last_err = sealer.get_last_error()
    if last_err:
        log.warning("Last error: %s", last_err)

    # ── 9. Disconnect ───────────────────────────────────────────────
    sealer.close()
    log.info("Done -- PlateLoc disconnected.")


if __name__ == "__main__":
    main()
