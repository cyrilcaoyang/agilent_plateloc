"""Service layer that exposes the PlateLoc driver as a spec-compliant
`EquipmentStatus` source.

Why this exists
---------------
The driver in ``plateloc.py`` is a thin wrapper around the ActiveX COM
control. It is synchronous and single-threaded: only one caller may
talk to the COM object at a time. The dashboard, however, polls
``GET /status`` every 2-3 seconds while operators may concurrently fire
``POST /control/*`` commands.

The service owns:

* a single driver instance (real or in-memory stub),
* an ``asyncio.Lock`` that serialises every call into the driver,
* a small in-memory state machine (``_busy_state``, ``_last_error``)
  used to compute the spec ``equipment_status`` field,
* a ``get_status()`` method that produces a fresh ``EquipmentStatus``
  envelope without ever issuing a write to the device.

If the real driver cannot be loaded (non-Windows host, missing ActiveX,
hardware off) ``dry_run=True`` swaps in a stub so the API surface stays
identical and the dashboard can be developed end-to-end on macOS/Linux.
"""

from __future__ import annotations

import asyncio
import logging
import socket
import time
from datetime import datetime, timezone
from typing import Any, Awaitable, Callable

from . import config as _config
from .models import (
    PROTOCOL_VERSION,
    ComponentStatus,
    EquipmentStatus,
    ErrorInfo,
    MetricValue,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Stub driver for dry-run / non-Windows development
# ---------------------------------------------------------------------------


class _StubPlateLoc:
    """In-memory mock that mirrors the public ``PlateLoc`` surface.

    Only the methods the service touches are implemented; anything else
    will raise ``AttributeError`` if accidentally used.
    """

    def __init__(self) -> None:
        self.com_port = "DRY-RUN"
        self._connected = False
        self._set_temp = 170
        self._set_time = 1.2
        self._actual_temp = 22  # ambient
        self._cycle_count = 0

    # lifecycle
    def connect(self, profile: str | None = None) -> None:  # noqa: ARG002
        self._connected = True
        self._actual_temp = self._set_temp  # heat up instantly

    def close(self) -> None:
        self._connected = False

    # control
    def set_sealing_temperature(self, t: int) -> int:
        self._set_temp = int(t)
        self._actual_temp = self._set_temp
        return 0

    def set_sealing_time(self, s: float) -> int:
        self._set_time = float(s)
        return 0

    def start_cycle(self) -> int:
        return 0

    def stop_cycle(self) -> int:
        self._cycle_count += 1
        return 0

    def move_stage_in(self) -> int:
        return 0

    def move_stage_out(self) -> int:
        return 0

    # readings
    def get_actual_temperature(self) -> int:
        return self._actual_temp

    def get_sealing_temperature(self) -> int:
        return self._set_temp

    def get_sealing_time(self) -> float:
        return self._set_time

    def get_cycle_count(self) -> int:
        return self._cycle_count

    def get_firmware_version(self) -> str:
        return "DRY-RUN-1.0"

    def get_version(self) -> str:
        return "DRY-RUN-AX"

    def enumerate_profiles(self) -> list[str]:
        return ["dry_run_default"]


# ---------------------------------------------------------------------------
# Service
# ---------------------------------------------------------------------------


_RECENT_ERROR_WINDOW_S = 60.0  # how long an error keeps the device in `error`


class PlateLocService:
    """Wraps a ``PlateLoc`` (or ``_StubPlateLoc``) driver and produces
    spec-compliant ``EquipmentStatus`` snapshots.

    Concurrency: all driver I/O happens inside ``self._lock``. Status
    reads share the same lock so a poll cannot interleave with a write.
    """

    def __init__(
        self,
        dry_run: bool = False,
        *,
        driver_factory: Callable[[], Any] | None = None,
    ) -> None:
        """
        Parameters
        ----------
        dry_run:
            When True the service uses ``_StubPlateLoc`` and reports
            ``equipment_status: dry_run`` regardless of operation.
        driver_factory:
            Optional override that returns a driver instance. Tests use
            this to inject a stub while keeping ``dry_run=False`` so
            the operational state machine (ready/busy/error) is exercised.
        """
        self.dry_run = dry_run
        self._driver_factory = driver_factory
        self._driver: Any | None = None
        self._lock = asyncio.Lock()
        self._started_at = time.monotonic()
        self._last_error: ErrorInfo | None = None
        self._busy_state: bool = False
        self._connect_profile: str | None = None

        # Identity (configurable so a deployment can override).
        self.equipment_id: str = _config.get("dashboard", "equipment_id", "plateloc")
        self.equipment_name: str = _config.get(
            "dashboard", "equipment_name", "Agilent PlateLoc"
        )
        self.equipment_kind = "plate_sealer"
        self.equipment_version: str | None = _config.get(
            "dashboard", "equipment_version", None
        )

    # ---- lifecycle ---------------------------------------------------------

    def _create_driver(self) -> Any:
        if self._driver_factory is not None:
            return self._driver_factory()
        if self.dry_run:
            return _StubPlateLoc()
        # Imported lazily so non-Windows hosts can run the dry-run service
        # without pywin32 installed.
        from .plateloc import PlateLoc

        return PlateLoc()

    async def startup(self, profile: str | None = None) -> None:
        """Create (or reuse) the driver and connect.

        On failure, leaves the service in `requires_init` and re-raises
        so callers (lifespan / `/control/startup`) can decide whether to
        log-and-continue or surface a 503.
        """
        async with self._lock:
            if self._driver is not None and self._driver_connected():
                return
            self._driver = self._create_driver()
            self._connect_profile = profile
            try:
                await asyncio.to_thread(self._driver.connect, profile)
                self._last_error = None
            except Exception as exc:
                self._record_error(exc, "startup")
                # keep self._driver around so retries reuse the same instance
                raise

    async def shutdown(self) -> None:
        """Best-effort disconnect. Never raises."""
        async with self._lock:
            if self._driver is None:
                return
            try:
                await asyncio.to_thread(self._driver.close)
            except Exception:
                logger.exception("Error while closing driver")
            finally:
                self._driver = None
                self._busy_state = False

    # ---- control -----------------------------------------------------------

    async def set_sealing_temperature(self, t: int) -> None:
        await self._do(
            "set_sealing_temperature",
            lambda d: d.set_sealing_temperature(int(t)),
        )

    async def set_sealing_time(self, s: float) -> None:
        await self._do(
            "set_sealing_time",
            lambda d: d.set_sealing_time(float(s)),
        )

    async def start_cycle(self) -> None:
        await self._do("start_cycle", lambda d: d.start_cycle())
        self._busy_state = True

    async def stop_cycle(self) -> None:
        await self._do("stop_cycle", lambda d: d.stop_cycle())
        self._busy_state = False

    async def move_stage_in(self) -> None:
        await self._do("move_stage_in", lambda d: d.move_stage_in())

    async def move_stage_out(self) -> None:
        await self._do("move_stage_out", lambda d: d.move_stage_out())

    async def _do(self, name: str, fn: Callable[[Any], Any]) -> None:
        async with self._lock:
            if self._driver is None or not self._driver_connected():
                raise RuntimeError(
                    "PlateLoc is not connected. POST /control/startup first."
                )
            try:
                await asyncio.to_thread(fn, self._driver)
                self._last_error = None
            except Exception as exc:
                self._record_error(exc, name)
                raise

    # ---- status (side-effect-free) ----------------------------------------

    async def get_status(self) -> EquipmentStatus:
        """Produce a fresh status snapshot. MUST NOT mutate hardware state.

        The spec requires this endpoint to be safe to call every 2-3
        seconds and to always return HTTP 200 unless the process itself
        is broken. We therefore catch every per-getter failure and fold
        it into ``equipment_status: degraded`` rather than raising.
        """
        async with self._lock:
            return self._build_status()

    def _build_status(self) -> EquipmentStatus:
        now = datetime.now(timezone.utc)
        uptime = time.monotonic() - self._started_at
        host = socket.gethostname()

        # ---- not connected: requires_init --------------------------------
        if self._driver is None or not self._driver_connected():
            return EquipmentStatus(
                protocol_version=PROTOCOL_VERSION,
                equipment_id=self.equipment_id,
                equipment_name=self.equipment_name,
                equipment_kind=self.equipment_kind,  # type: ignore[arg-type]
                equipment_version=self.equipment_version,
                host=host,
                equipment_status="requires_init",
                message="Driver not connected. POST /control/startup to initialize.",
                required_actions=["startup"],
                device_time=now,
                uptime_seconds=uptime,
                last_error=self._last_error,
            )

        # ---- read what we can; never let a single getter fail status -----
        metrics: dict[str, MetricValue] = {}
        details: dict[str, Any] = {}
        readback_errors: list[str] = []

        def _read(label: str, fn: Callable[[], Any]) -> Any:
            try:
                return fn()
            except Exception as exc:
                readback_errors.append(f"{label}: {exc}")
                return None

        actual_temp = _read("actual_temperature", self._driver.get_actual_temperature)
        if actual_temp is not None:
            metrics["actual_temperature"] = MetricValue(
                value=actual_temp, unit="C", timestamp=now
            )
        setpoint = _read("setpoint_temperature", self._driver.get_sealing_temperature)
        if setpoint is not None:
            metrics["setpoint_temperature"] = MetricValue(
                value=setpoint, unit="C", timestamp=now
            )
        seal_time = _read("sealing_time", self._driver.get_sealing_time)
        if seal_time is not None:
            metrics["sealing_time"] = MetricValue(
                value=seal_time, unit="s", timestamp=now
            )
        cycle_count = _read("cycle_count", self._driver.get_cycle_count)
        if cycle_count is not None:
            metrics["cycle_count"] = MetricValue(value=cycle_count, unit="count")

        firmware = _read("firmware_version", self._driver.get_firmware_version)
        if firmware:
            details["firmware_version"] = firmware
        ax_version = _read("activex_version", self._driver.get_version)
        if ax_version:
            details["activex_version"] = ax_version
        if self._connect_profile:
            details["profile"] = self._connect_profile
        com_port = getattr(self._driver, "com_port", None)
        if com_port:
            details["com_port"] = com_port

        # ---- components --------------------------------------------------
        connected = self._driver_connected()
        sealer_state = (
            "busy" if self._busy_state else ("idle" if connected else "disconnected")
        )
        components: dict[str, ComponentStatus] = {
            "sealer": ComponentStatus(
                connected=connected,
                state=sealer_state,
            ),
            # The ActiveX control does not expose a stage position query,
            # so we report it as `unknown` until events are wired in.
            "stage": ComponentStatus(connected=connected, state="unknown"),
        }

        # ---- top-level equipment_status ----------------------------------
        if self.dry_run:
            state: str = "dry_run"
            message: str | None = "Dry-run mode - no hardware connected"
            details["dry_run"] = True
        elif self._busy_state:
            state = "busy"
            message = "Seal cycle in progress"
        elif self._last_error is not None and (
            (now - self._last_error.timestamp).total_seconds()
            < _RECENT_ERROR_WINDOW_S
        ):
            state = "error"
            message = self._last_error.message
        elif readback_errors:
            state = "degraded"
            message = "; ".join(readback_errors)
        else:
            state = "ready"
            message = "Idle, ready to seal"

        return EquipmentStatus(
            protocol_version=PROTOCOL_VERSION,
            equipment_id=self.equipment_id,
            equipment_name=self.equipment_name,
            equipment_kind=self.equipment_kind,  # type: ignore[arg-type]
            equipment_version=self.equipment_version,
            host=host,
            equipment_status=state,  # type: ignore[arg-type]
            message=message,
            device_time=now,
            uptime_seconds=uptime,
            components=components,
            metrics=metrics,
            last_error=self._last_error,
            details=details,
        )

    # ---- helpers -----------------------------------------------------------

    def _driver_connected(self) -> bool:
        """Driver is connected if either flag is set. The real PlateLoc
        uses the private `_connected` attribute; the stub also exposes it
        for parity. Wrapped in getattr so an unexpected driver type
        cannot crash the status endpoint."""
        if self._driver is None:
            return False
        return bool(getattr(self._driver, "_connected", False))

    def _record_error(self, exc: Exception, code: str) -> None:
        self._last_error = ErrorInfo(
            code=code,
            message=str(exc),
            severity="error",
            timestamp=datetime.now(timezone.utc),
        )
        logger.exception("PlateLoc error in %s", code)


__all__ = ["PlateLocService", "_StubPlateLoc"]
