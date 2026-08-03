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

from . import __version__
from . import config as _config
from .claims import ClaimStore
from .models import (
    PROTOCOL_VERSION,
    ComponentStatus,
    EquipmentStatus,
    ErrorInfo,
    MetricValue,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# allowed_actions — one pure function, consulted by both surfaces
#
# Skill names mirror the SDK skill catalog for `kind=plate_sealer` (see
# lab_skills/skill_catalog/plate_sealer.py). The device is the source of
# truth: the SDK prefers our allowed_actions over its own catalog
# `requires_states` whenever the field is non-empty.
#
# §6.2 single-source-of-truth: the seal.start entry is gated by the SAME
# `evaluate_*_interlock` helpers the `/control/seal/start` 412 path uses, so
# the advertised list and the refusals cannot drift.
#
# v1.2 (§2.3): the list must agree with `activity` as well as with the
# top-level state. While a seal cycle is executing, everything that would
# start a *second* run or move the carriage is withheld; the abort/stop class
# stays listed so an abort is always reachable.
# ---------------------------------------------------------------------------

_ALL_PLATE_SEALER_SKILLS = [
    "startup",
    "shutdown",
    "seal.start",
    "seal.stop",
    "seal.set_temperature",
    "seal.set_time",
    "stage.in",
    "stage.out",
]

#: Offered while the device is initialized, healthy and idle, in catalog
#: order. `seal.stop` is deliberately absent — there is nothing to stop.
_IDLE_SKILLS = [
    "startup",
    "shutdown",
    "seal.start",
    "seal.set_temperature",
    "seal.set_time",
    "stage.in",
    "stage.out",
]

#: Offered while `activity == "running"`: abort/stop class only (§2.3).
_RUNNING_SKILLS = ["shutdown", "seal.stop"]


def _compute_allowed_actions(
    state: str,
    activity: str,
    *,
    stage_state: str,
    seal_start_blocked: bool,
) -> list[str]:
    """The device's authoritative "what would I honor right now" list.

    Rules:

    * ``requires_init`` → only ``startup``.
    * ``e_stop`` / ``unknown`` → nothing.
    * ``activity == "running"`` → no second concurrent cycle and no
      carriage move while the press is down: ``shutdown`` + ``seal.stop``
      (§2.3). Checked *before* the health branches so a fault that lands
      mid-cycle cannot take the abort action away.
    * otherwise the idle set, minus ``seal.start`` when any of the three §6
      interlocks would refuse it, minus the no-op stage direction.

    ``error`` and ``degraded`` deliberately do **not** collapse to
    ``["shutdown"]`` (which is what this device advertised through v1.3.2).
    Two reasons. §2.2 says the run-blocking fault must remove the *run*,
    while "recovery, abort, standby, and diagnostic actions may remain
    available when safe" — and after a failed cycle the operator's recovery
    is exactly ``stage.out`` (retrieving the plate from a hot chamber), which
    the old list withheld while the endpoint honored it anyway. That
    mismatch was also a §6.2 violation: ``/status`` omitted actions the
    device would have performed. The run itself is gated instead by
    :meth:`PlateLocService.evaluate_health_interlock`, which the
    ``/control/seal/start`` 412 path shares.
    """
    if state == "requires_init":
        return ["startup"]
    if state in ("e_stop", "unknown"):
        return []
    if activity == "running":
        return list(_RUNNING_SKILLS)

    allowed = [
        skill
        for skill in _IDLE_SKILLS
        if not (skill == "seal.start" and seal_start_blocked)
    ]
    # Stage move dedup: don't advertise the no-op direction. A POST to the
    # "already there" direction is still accepted (the device treats it as a
    # 200 no-op); we just leave it out so an operator UI doesn't render a
    # redundant button. Asymmetry vs seal.start is deliberate: a redundant
    # stage move is harmless, sealing without a plate wastes hot air.
    if stage_state == "in":
        allowed = [s for s in allowed if s != "stage.in"]
    elif stage_state == "out":
        allowed = [s for s in allowed if s != "stage.out"]
    return allowed


def _coerce_float(value: Any) -> float | None:
    """Best-effort cast to ``float`` for ActiveX values that can be ``None``,
    ``int``, ``float``, or a parseable string. Returns ``None`` if the
    value cannot be coerced — callers treat that as "temperature
    unavailable" and fail closed."""
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


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
        # The real control runs blocking (``PlateLoc(blocking=True)``): the
        # call returns when the physical cycle is done, and the instrument's
        # odometer has advanced by one. The stub mirrors both.
        self._cycle_count += 1
        return 0

    def stop_cycle(self) -> int:
        # Stopping is not a completed cycle — the odometer does not move.
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


# ---------------------------------------------------------------------------
# last_error.code taxonomy (v1.3.1)
#
# A closed set of identifiers for *what* failed, separate from `message`
# (the driver's free-form text). The dashboard renders a targeted
# recovery hint by branching on `code`; it MUST NOT regex on `message`.
#
# `_classify_error` is the single classifier: given the failing method
# name, the exception, and the driver's `get_last_error()` detail, it
# returns one of these codes. Order in the classifier matters — text
# matches (low_air_pressure, heater_*) win over context fallbacks
# (stage_jam, com_init_failed) so a specific cause beats "the stage
# move endpoint failed for some reason".
#
# Keep this list minimal. Add a code only when the codebase actually
# raises that mode AND the dashboard would render differently for it.
# "com_other" is the catch-all; falling into it on a previously-seen
# failure mode is the signal to add a new code, not to grow this comment.
# ---------------------------------------------------------------------------

LAST_ERROR_CODES: frozenset[str] = frozenset(
    {
        "low_air_pressure",
        "no_plate",
        "vacuum_error",
        "com_init_failed",
        "com_timeout",
        "com_other",
        "heater_overtemp",
        "heater_undertemp",
        "profile_not_found",
        "stage_jam",
        "process_internal",
    }
)


def _classify_error(
    method_name: str, exc: Exception, detail: str | None
) -> str:
    """Return one of ``LAST_ERROR_CODES`` for the given failure.

    Order is deliberate:

    1. ``process_internal`` — Python type errors (KeyError, etc.)
       indicate a software bug, not a driver fault. Distinguishing
       these is the dashboard's signal to file a ticket rather than
       reach for the diagnostics dialog.
    2. Specific text matches — ``low_air_pressure``, ``no_plate``,
       ``vacuum_error``, ``heater_*``, ``profile_not_found``. Driver text
       is the most reliable signal.
    3. ``com_timeout`` — timeout substring beats generic com_other.
    4. Context fallbacks — ``stage_jam`` if the failing method was a
       stage move; ``com_init_failed`` if the failing method was
       startup. These fire when the driver text is unhelpful (empty
       or generic HRESULT).
    5. ``com_other`` — default; anything we don't yet classify.
    """
    if isinstance(exc, (KeyError, AttributeError, TypeError, NameError)):
        return "process_internal"
    return _classify_error_text(
        method_name,
        f"{detail or ''} {exc}",
        is_timeout=isinstance(exc, TimeoutError),
    )


def _classify_error_text(
    method_name: str, text: str, *, is_timeout: bool = False
) -> str:
    """Classify a fault from its message text (steps 2-5 of
    :func:`_classify_error`).

    Split out so a *readback* fault observed while composing ``/status`` —
    which has a message but no exception object — lands on the same stable
    taxonomy as an operational failure, instead of reaching clients only as
    free text they would have to regex (best-practice #6).
    """
    haystack = text.lower()

    # Profile mis-config: only meaningful at startup time. The
    # PlateLoc driver wraps Initialize() failures with the available
    # profile list in the exception text — match on that, not on the
    # generic HRESULT.
    if method_name == "startup" and (
        "profile" in haystack
        and ("not found" in haystack or "available profiles" in haystack)
    ):
        return "profile_not_found"

    # Driver-text matches — most specific signals.
    if "air pressure" in haystack:
        return "low_air_pressure"
    # Seal-cycle physical faults. Distinct, actionable driver strings that
    # otherwise fell through to com_other (observed live 2026-07-15).
    if "no plate" in haystack:
        return "no_plate"
    if "vacuum" in haystack:
        return "vacuum_error"
    if "overtemp" in haystack or "over temp" in haystack or "over-temperature" in haystack:
        return "heater_overtemp"
    if (
        "undertemp" in haystack
        or "under temp" in haystack
        or "did not reach setpoint" in haystack
    ):
        return "heater_undertemp"

    # Timeouts: TimeoutError type OR "timeout"/"timed out" substring.
    if is_timeout or "timeout" in haystack or "timed out" in haystack:
        return "com_timeout"

    # Context fallbacks — used when the text alone isn't decisive.
    if method_name in ("move_stage_in", "move_stage_out"):
        return "stage_jam"
    if method_name == "startup":
        return "com_init_failed"

    return "com_other"


def _readback_error_info(
    readback_errors: list[str], now: datetime
) -> ErrorInfo | None:
    """Synthesize ``last_error`` from an active readback fault.

    A failed instrument readback is a real, diagnosable fault, but it is not
    an *operational* failure — nothing was executing, so it never touched
    ``self._last_error`` and reached ``/status`` only as free text in
    ``message``. Identifying it meant string-matching that text, exactly what
    best-practice #6 warns against.

    Severity is ``warning``, not ``error``: §2.2 already carries the safety
    consequence by putting the top-level state at ``degraded``, and a useful
    subset of capability remains. This mirrors the reference shaker envelope
    in STATUS_SPEC §10.
    """
    if not readback_errors:
        return None
    # Readback strings are "<label>: <exception>" (see _read_driver_metrics).
    _, _, detail = readback_errors[0].partition(": ")
    return ErrorInfo(
        code=_classify_error_text("status_readback", detail or readback_errors[0]),
        message="; ".join(readback_errors),
        severity="warning",
        timestamp=now,
    )


class TemperatureOutOfBand(Exception):
    """Raised by ``start_cycle`` when the heater is not at setpoint.

    Layer-1 interlock from ``docs/INTERLOCKS.md``: the device refuses to
    start a seal cycle when ``abs(actual - setpoint) > tolerance``.
    Carries enough structured data for the API layer to emit a clean
    HTTP 412 body without re-querying the driver.
    """

    def __init__(
        self,
        message: str,
        *,
        actual_c: float | None,
        setpoint_c: float | None,
        tolerance_c: float,
        retry_after_s: float | None,
    ) -> None:
        super().__init__(message)
        self.actual_c = actual_c
        self.setpoint_c = setpoint_c
        self.tolerance_c = tolerance_c
        self.retry_after_s = retry_after_s


class RecentFailureNotCleared(Exception):
    """Raised by ``start_cycle`` while an operational failure is still recent.

    Layer-1 interlock (v1.4.0), the third of three and the one §2.2 asks for
    directly: do not start a normal run while the device knows of an active
    fault. It gates ``seal.start`` **only** — recovery, abort and diagnostic
    actions stay available, which is what lets an operator drive
    ``stage.out`` after a mid-cycle air failure instead of being offered
    nothing but ``shutdown``.

    Clears exactly as ``last_error`` does (§6.4): the first 2xx from any
    operational endpoint drops it, and the window expires on its own.
    """

    def __init__(
        self,
        message: str,
        *,
        last_error_code: str | None,
        last_error_message: str,
        retry_after_s: float | None,
    ) -> None:
        super().__init__(message)
        self.last_error_code = last_error_code
        self.last_error_message = last_error_message
        self.retry_after_s = retry_after_s


class StageNotLoaded(Exception):
    """Raised by ``start_cycle`` when the plate stage is not in the
    loaded ("in") position.

    Layer-1 interlock (v1.3.0): the device refuses to start a seal
    cycle unless ``components.stage.state == "in"``. The Agilent COM
    API does not expose a stage-position query, so the position is
    command-tracked and resets to ``"unknown"`` on process restart
    or mid-cycle failure (the plate may have been moved manually
    while the service was down — refusing safely beats guessing).

    Carries the stage state at refusal time so the API layer can
    emit ``{"detail", "stage_state", "required": "in"}`` without
    re-reading service state.
    """

    def __init__(self, message: str, *, stage_state: str) -> None:
        super().__init__(message)
        self.stage_state = stage_state


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
        enforce_claims: bool = True,
        enforce_temp_interlock: bool = True,
        enforce_stage_interlock: bool = True,
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
        enforce_claims:
            STATUS_SPEC v1.1 strictness switch. When True (default), the
            API layer rejects ``/control/*`` requests with HTTP 423 unless
            they carry a valid ``X-Claim-Token``. Set False for the
            handful of legacy / single-operator deployments that want
            v1.1 *advisory* claims (the device still publishes
            ``allowed_actions`` and ``details.claimed_by`` but does not
            block writes from clients without a token).
        enforce_temp_interlock:
            Layer-1 interlock (see ``docs/INTERLOCKS.md``). When True
            (default), ``start_cycle`` raises ``TemperatureOutOfBand`` if
            the heater is not within ``temperature_tolerance_c`` of the
            setpoint, which the API surfaces as HTTP 412. Set False only
            for emergency overrides (e.g. calibration at room
            temperature) — running with this off restores the failure
            mode where sealing below setpoint produces an underspec'd
            seal and downstream pneumatic faults.
        enforce_stage_interlock:
            Layer-1 interlock (v1.3.0). When True (default),
            ``start_cycle`` raises ``StageNotLoaded`` if the stage is
            not in the loaded position, which the API surfaces as
            HTTP 412 with ``{"detail":"Stage not loaded","stage_state":
            "out"|"unknown","required":"in"}``. Independent of
            ``enforce_temp_interlock``. Set False only for emergency
            overrides — running with this off restores the failure
            mode where seal cycles run with the carriage extended,
            wasting hot air and risking film damage.
        """
        self.dry_run = dry_run
        self._driver_factory = driver_factory
        self._driver: Any | None = None
        # Two locks, ordered state -> io (nothing takes the state lock while
        # holding io):
        #
        # * ``_lock`` — state lock. Guards the in-memory bookkeeping
        #   (driver presence, busy/stage/activity, last_error). Never held
        #   across a slow COM transaction, so a ``/status`` poll cannot
        #   queue behind a seal cycle. That is what makes v1.2 ``activity``
        #   observable at all: ``StartCycle`` blocks for the whole cycle,
        #   and a reader that had to wait for the state lock would only ever
        #   see the device before or after it, never `running`.
        # * ``_io_lock`` — COM channel lock. Held around every instrument
        #   transaction: the ActiveX control is single-threaded and the
        #   32-bit surrogate serves one request at a time over its pipe.
        self._lock = asyncio.Lock()
        self._io_lock = asyncio.Lock()
        self._started_at = time.monotonic()
        self._last_error: ErrorInfo | None = None
        self._busy_state: bool = False
        self._connect_profile: str | None = None
        # Activity span tracking (STATUS_SPEC v1.2 §2.3). ``_activity`` is
        # the last observed value; ``_activity_since`` is the instant it last
        # changed. Both are stamped by the methods that own the transition
        # (start_cycle / stop_cycle / startup / shutdown); ``_compose_status``
        # only reconciles.
        self._activity: str = "unknown"
        self._activity_since: datetime | None = None
        self._cycle_started_at: datetime | None = None
        # Last successful instrument readback, reused while a seal cycle owns
        # the COM channel (see ``get_status``).
        self._readings: dict[str, Any] = {}
        self._readback_errors: list[str] = []
        self._readings_at: datetime | None = None
        # Stage position is command-tracked (the COM API has no
        # GetStagePosition equivalent). Defaults to "unknown" at process
        # start; the operator homes it via /control/stage/{in,out}. See
        # README "Stage interlock" for the full transition table.
        self._stage_state: str = "unknown"
        self.enforce_claims = enforce_claims
        self.enforce_temp_interlock = enforce_temp_interlock
        self.enforce_stage_interlock = enforce_stage_interlock
        self.claims = ClaimStore()

        # Tolerance (in C) inside which `actual_temperature` is considered
        # to have reached `setpoint_temperature`. Mirrors what `demo.py`
        # uses for its temperature-wait loop, so the device speaks the
        # same language as the operator-facing demo. The ActiveX control
        # does not expose a native "temperature stable" signal; we
        # synthesize it by comparing the two metrics.
        self._temp_tolerance_c: float = float(
            _config.get("film", "temperature_tolerance_c", 2)
        )

        # Identity (configurable so a deployment can override).
        self.equipment_id: str = _config.get("dashboard", "equipment_id", "plateloc")
        self.equipment_name: str = _config.get(
            "dashboard", "equipment_name", "Agilent PlateLoc"
        )
        self.equipment_kind = "plate_sealer"
        # Fall back to the package version rather than publishing null: an
        # unset `[dashboard] equipment_version` should not cost the dashboard
        # the ability to tell which build the device is running.
        self.equipment_version: str | None = (
            _config.get("dashboard", "equipment_version", None) or __version__
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

    @property
    def connected(self) -> bool:
        """True once the driver is attached and reports connected.

        A failed ``startup`` keeps ``self._driver`` around for retries,
        so driver presence alone is not connection.
        """
        return self._driver_connected()

    async def startup(self, profile: str | None = None) -> None:
        """Create (or reuse) the driver and connect.

        On failure, leaves the service in `requires_init` and re-raises
        so callers (lifespan / `/control/startup`) can decide whether to
        log-and-continue or surface a 503.

        Does NOT clear ``self._last_error`` on success: the API layer
        owns that policy and clears only when the overall endpoint
        response is 2xx (see :meth:`clear_last_error_on_success`).
        """
        async with self._lock:
            if self._driver is not None and self._driver_connected():
                return
            self._driver = self._create_driver()
            self._connect_profile = profile
            try:
                await self._io(self._driver.connect, profile)
            except Exception as exc:
                # `connect()` already calls get_last_error() and folds the
                # detail into the exception text on the Initialize-failed
                # path, but other startup failures (driver-create, ATL
                # hosting, surrogate exit) won't carry that string. Best-
                # effort enrich here too.
                detail = await self._read_driver_last_error()
                self._record_error(exc, "startup", detail=detail)
                # keep self._driver around so retries reuse the same instance
                raise
            self._invalidate_readings()
            # A freshly connected sealer is not cycling (§2.3 pins
            # requires_init ⇒ idle; this is the transition out of it).
            self._note_activity("idle")

    async def shutdown(self) -> None:
        """Best-effort disconnect. Never raises.

        Resets ``_stage_state`` to ``"unknown"`` — the next operator
        cycle must re-home the carriage. Does NOT clear
        ``self._last_error`` (the API layer owns that policy; see
        :meth:`clear_last_error_on_success`).
        """
        async with self._lock:
            if self._driver is None:
                # Even on the no-op path, stage state should reflect
                # "we cannot vouch for the carriage position" — same
                # rationale as a fresh process start.
                self._stage_state = "unknown"
                self._note_activity("idle")
                return
            try:
                await self._io(self._driver.close)
            except Exception:
                logger.exception("Error while closing driver")
            finally:
                self._driver = None
                self._busy_state = False
                self._cycle_started_at = None
                self._stage_state = "unknown"
                self._invalidate_readings()
                # Disconnected hardware cannot be sealing under our control;
                # §2.3 pins requires_init ⇒ idle.
                self._note_activity("idle")

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
        """Run one seal cycle.

        The ActiveX control is configured in **blocking** mode
        (``PlateLoc(blocking=True)``), so ``StartCycle`` returns only when
        the physical cycle has finished. The seal cycle — this device's
        primary operation — is therefore exactly the span of that COM call,
        and that is the span v1.2 reports as ``activity: "running"`` (§2.3).
        Two consequences shape the code below:

        * The state lock is **not** held across the call. It used to be,
          which meant a ``/status`` poll queued behind the whole cycle and
          no reader could ever observe the device while it was sealing.
        * ``_busy_state`` is set before the call and cleared after it (on
          both the success and failure paths), instead of being latched on
          afterwards until an explicit ``/control/seal/stop``.

        Two layer-1 interlocks fire before the hardware moves:

        1. **Stage** — raises :class:`StageNotLoaded` unless
           ``components.stage.state == "in"``. Checked first because
           a wrong-stage refusal is faster for the operator to fix
           (a single click) than a temperature ramp.
        2. **Temperature** — raises :class:`TemperatureOutOfBand` if
           the heater is outside ``temperature_tolerance_c`` of the
           setpoint.

        Both pre-flight checks leave hardware state untouched, so
        ``_stage_state`` is unchanged on either refusal path.

        On the COM path: the cycle physically commits the stage to
        ``"in"`` (the carriage is under the press at the start and
        stays there at the end). We pessimize to ``"unknown"`` on
        entry so a mid-cycle failure (driver fault after the
        physical commit started) leaves a truthful state instead of
        a stale ``"in"``.

        Raises
        ------
        RuntimeError
            Driver not connected, or a cycle is already in flight. Both are
            device-state conflicts (HTTP 409), not precondition refusals —
            see STATUS_SPEC §6.1 on 409-vs-412.
        StageNotLoaded, RecentFailureNotCleared, TemperatureOutOfBand
            Layer-1 interlock refusals (HTTP 412), checked in that order:
            the two in-memory gates before the one that needs COM reads.
        """
        async with self._lock:
            driver = self._driver
            if driver is None or not self._driver_connected():
                raise RuntimeError(
                    "PlateLoc is not connected. POST /control/startup first."
                )
            self._assert_not_busy()
            # In-memory gates first (cheap): stage position, then an
            # uncleared recent failure.
            self._assert_stage_loaded()
            self._assert_failure_cleared(self._last_error)

        # Temperature gate second: it needs COM reads, which must not run
        # under the state lock. Both gates use the same
        # ``evaluate_*_interlock`` helpers that build allowed_actions, so the
        # two surfaces agree (§6.2).
        await self._assert_temperature_in_band(driver)

        async with self._lock:
            # Re-check after the reads: the driver may have gone away and a
            # racing caller may have taken the cycle in the meantime.
            if self._driver is not driver or not self._driver_connected():
                raise RuntimeError(
                    "PlateLoc disconnected during seal-cycle pre-flight."
                )
            self._assert_not_busy()
            self._assert_stage_loaded()
            self._assert_failure_cleared(self._last_error)
            # Every pre-flight passed; from here on the COM call may
            # mutate physical stage position. Pessimize.
            self._stage_state = "unknown"
            self._busy_state = True
            self._cycle_started_at = datetime.now(timezone.utc)
            self._note_activity("running")  # exact span start (§2.3)

        try:
            await self._io(driver.start_cycle)
        except Exception as exc:
            detail = await self._read_driver_last_error()
            async with self._lock:
                self._busy_state = False
                self._cycle_started_at = None
                self._note_activity("idle")
                self._invalidate_readings()
                self._record_error(exc, "start_cycle", detail=detail)
                # Leave _stage_state as "unknown" — the cycle aborted
                # mid-motion and the carriage position is no longer
                # tracked.
            raise

        async with self._lock:
            self._busy_state = False
            self._cycle_started_at = None
            self._note_activity("idle")
            self._stage_state = "in"
            # The instrument's odometer just advanced; drop the cached
            # readback so the next poll reports the new cycles_total.
            self._invalidate_readings()

    def evaluate_temperature_interlock(
        self,
        actual: float | None,
        setpoint: float | None,
    ) -> tuple[bool, dict[str, Any] | None]:
        """Single source of truth for the layer-1 temperature interlock.

        Returns ``(should_block, body_for_412)``:

        * ``should_block`` is ``False`` when the heater is at setpoint
          within ``self._temp_tolerance_c`` (the seal cycle would be
          honoured) **or** when ``enforce_temp_interlock`` is disabled.
          In both cases ``body_for_412`` is ``None``.
        * ``should_block`` is ``True`` when the band check fails or
          when the temperatures cannot be read (fail-closed). The
          returned ``body_for_412`` is the structured JSON body that
          ``/control/seal/start`` returns with HTTP 412 — building it
          here keeps the ``/status`` ``allowed_actions`` gate and the
          412 refusal path on the same single answer.
        """
        if not self.enforce_temp_interlock:
            return False, None

        tolerance = self._temp_tolerance_c

        if actual is None or setpoint is None:
            return True, {
                "detail": "Cannot verify temperature: actual or setpoint unavailable",
                "actual_c": actual,
                "setpoint_c": setpoint,
                "tolerance_c": tolerance,
                "retry_after_s": None,
            }

        delta = actual - setpoint
        if abs(delta) <= tolerance:
            return False, None

        # Best-effort retry estimate. PlateLoc heat-up is faster than
        # cool-down on the hot-plate; the constants below are
        # conservative averages, not measured ramps, and intentionally
        # round up so callers don't hot-poll. None is also a valid
        # answer; we always provide a number here so the dashboard can
        # render "try again in ~N s".
        excess = abs(delta) - tolerance
        ramp_c_per_s = 1.0 if delta < 0 else 0.3
        retry_after_s = max(1.0, round(excess / ramp_c_per_s + 0.5))

        return True, {
            "detail": "Temperature outside seal band",
            "actual_c": actual,
            "setpoint_c": setpoint,
            "tolerance_c": tolerance,
            "retry_after_s": retry_after_s,
        }

    def evaluate_stage_interlock(
        self,
        stage_state: str,
    ) -> tuple[bool, dict[str, Any] | None]:
        """Single source of truth for the stage-position interlock.

        Returns ``(should_block, body_for_412)``:

        * ``should_block`` is ``False`` when ``stage_state == "in"`` —
          the stage is loaded and a seal cycle can proceed — or when
          ``enforce_stage_interlock`` is disabled. In both cases
          ``body_for_412`` is ``None``.
        * ``should_block`` is ``True`` when ``stage_state`` is
          ``"out"`` or ``"unknown"``. The returned ``body_for_412``
          is the structured JSON body that ``/control/seal/start``
          returns with HTTP 412. No ``Retry-After`` — recovery is
          operator-driven (``POST /control/stage/in``), not
          time-based.

        Mirrors :meth:`evaluate_temperature_interlock` so the two
        interlocks compose uniformly at the seal.start endpoint and
        the /status allowed_actions builder.
        """
        if not self.enforce_stage_interlock:
            return False, None
        if stage_state == "in":
            return False, None
        return True, {
            "detail": "Stage not loaded",
            "stage_state": stage_state,
            "required": "in",
        }

    def evaluate_health_interlock(
        self,
        last_error: ErrorInfo | None,
    ) -> tuple[bool, dict[str, Any] | None]:
        """Single source of truth for the recent-failure interlock.

        Returns ``(should_block, body_for_412)``, the same contract as
        :meth:`evaluate_stage_interlock` and
        :meth:`evaluate_temperature_interlock`, so all three compose
        uniformly at ``/control/seal/start`` and in the ``/status``
        ``allowed_actions`` builder (§6.2).

        Blocks while ``last_error`` is inside the recent-error window — the
        same window that puts the device in ``equipment_status: "error"``, so
        the two surfaces cannot disagree about whether a run is available.
        Recovery is time-bounded (or immediate, via §6.4's auto-clear on the
        next successful action), so the body carries ``retry_after_s``.
        """
        if last_error is None:
            return False, None
        elapsed = (
            datetime.now(timezone.utc) - last_error.timestamp
        ).total_seconds()
        remaining = _RECENT_ERROR_WINDOW_S - elapsed
        if remaining <= 0:
            return False, None
        return True, {
            "detail": "Recent operational failure not cleared",
            "last_error_code": last_error.code,
            "last_error_message": last_error.message,
            "retry_after_s": max(1.0, round(remaining)),
        }

    def _assert_failure_cleared(self, last_error: ErrorInfo | None) -> None:
        """Raise :class:`RecentFailureNotCleared` if the health interlock
        would block a seal cycle right now.

        Caller MUST already hold ``self._lock``. Synchronous — the decision
        is in-memory.
        """
        blocks, body = self.evaluate_health_interlock(last_error)
        if not blocks:
            return
        assert body is not None  # invariant: blocks=True implies a body
        raise RecentFailureNotCleared(
            body["detail"],
            last_error_code=body["last_error_code"],
            last_error_message=body["last_error_message"],
            retry_after_s=body["retry_after_s"],
        )

    def _assert_not_busy(self) -> None:
        """Raise ``RuntimeError`` (→ HTTP 409) if a seal cycle is in flight.

        Mirrors the ``activity == "running"`` gate in
        :func:`_compute_allowed_actions`: while the device advertises only
        the abort/stop class, anything that would start a second run or move
        the carriage is refused. A concurrency conflict is a device-state
        conflict, so it keeps the 409 path rather than becoming a §6.1
        precondition 412.

        Caller MUST already hold ``self._lock``.
        """
        if self._busy_state:
            raise RuntimeError(
                "A seal cycle is in progress. Wait for it to finish "
                "(or POST /control/seal/stop)."
            )

    def _assert_stage_loaded(self) -> None:
        """Raise :class:`StageNotLoaded` if the stage interlock would
        block a seal cycle right now.

        Caller MUST already hold ``self._lock``. Synchronous because
        stage state is in-memory — no driver I/O needed.
        """
        blocks, body = self.evaluate_stage_interlock(self._stage_state)
        if not blocks:
            return
        assert body is not None  # invariant: blocks=True implies a body
        raise StageNotLoaded(body["detail"], stage_state=body["stage_state"])

    async def _assert_temperature_in_band(self, driver: Any) -> None:
        """Raise :class:`TemperatureOutOfBand` if the temperature
        interlock would block a seal cycle right now.

        Caller MUST NOT hold ``self._lock``: the two COM reads go through
        :meth:`_io` (worker thread, COM-channel lock), and holding the state
        lock across instrument I/O would stall ``/status`` polls. The
        decision itself is delegated to
        :meth:`evaluate_temperature_interlock` so this path stays in lockstep
        with the ``/status`` ``allowed_actions`` gate.
        """
        actual_raw = await self._io(driver.get_actual_temperature)
        setpoint_raw = await self._io(driver.get_sealing_temperature)

        blocks, body = self.evaluate_temperature_interlock(
            _coerce_float(actual_raw), _coerce_float(setpoint_raw)
        )
        if not blocks:
            return
        assert body is not None  # invariant: blocks=True implies a body
        raise TemperatureOutOfBand(
            body["detail"],
            actual_c=body["actual_c"],
            setpoint_c=body["setpoint_c"],
            tolerance_c=body["tolerance_c"],
            retry_after_s=body["retry_after_s"],
        )

    async def stop_cycle(self) -> None:
        """Stop the current seal cycle. Idempotent — stopping when nothing
        is running is a 2xx no-op.

        This is the abort-class action, so it is the one control call the
        device still honors while ``activity == "running"``. Note that with
        the control in blocking mode the in-flight ``StartCycle`` owns the
        COM channel: a stop issued mid-cycle is serialised behind it and
        lands as soon as that call returns.
        """
        await self._do(
            "stop_cycle", lambda d: d.stop_cycle(), allow_while_busy=True
        )
        async with self._lock:
            self._busy_state = False
            self._cycle_started_at = None
            self._note_activity("idle")
            self._invalidate_readings()

    async def move_stage_in(self) -> None:
        """Move the plate carriage to the loaded position.

        Tracks ``_stage_state`` per the v1.3.0 transition table.
        Inlined (not via :meth:`_do`) so the state mutations and the
        COM call sit inside the same critical section.
        """
        await self._move_stage("move_stage_in", "in")

    async def move_stage_out(self) -> None:
        """Move the plate carriage to the unloaded position.

        Tracks ``_stage_state`` per the v1.3.0 transition table.
        """
        await self._move_stage("move_stage_out", "out")

    async def _move_stage(self, com_method: str, target: str) -> None:
        """Shared implementation for ``move_stage_in`` and
        ``move_stage_out``. Pessimizes ``_stage_state`` on entry and
        commits to ``target`` only on a clean COM return.

        A POST /control/stage/{in,out} to the position the stage is
        already in is handled as a no-op 200 by the COM driver. The
        net state remains ``target``; a /status poll mid-call cannot
        observe the "unknown" flicker because /status takes the same
        lock.
        """
        async with self._lock:
            if self._driver is None or not self._driver_connected():
                raise RuntimeError(
                    "PlateLoc is not connected. POST /control/startup first."
                )
            # The carriage cannot move while the press is down (the
            # instrument itself reports "Stage cannot move - press is down").
            # Refuse up front so the advertised list and the refusals agree.
            self._assert_not_busy()
            self._stage_state = "unknown"
            try:
                await self._io(getattr(self._driver, com_method))
            except Exception as exc:
                detail = await self._read_driver_last_error()
                self._record_error(exc, com_method, detail=detail)
                # Leave _stage_state as "unknown" — the move failed
                # mid-motion and we no longer know where the carriage is.
                raise
            self._stage_state = target

    async def _do(
        self,
        name: str,
        fn: Callable[[Any], Any],
        *,
        allow_while_busy: bool = False,
    ) -> None:
        """Run one driver call under the state lock, recording failures.

        ``allow_while_busy`` is for the abort class only (``stop_cycle``):
        every other action is refused mid-cycle, matching the
        ``activity == "running"`` gate in :func:`_compute_allowed_actions`.
        """
        async with self._lock:
            if self._driver is None or not self._driver_connected():
                raise RuntimeError(
                    "PlateLoc is not connected. POST /control/startup first."
                )
            if not allow_while_busy:
                self._assert_not_busy()
            try:
                await self._io(fn, self._driver)
            except Exception as exc:
                # Best-effort: pull the human-readable message out of the
                # ActiveX control via GetLastError so operators see what
                # the instrument actually reported instead of just the
                # generic Agilent HRESULT (e.g. -2147221503 = 0x80040201).
                detail = await self._read_driver_last_error()
                self._record_error(exc, name, detail=detail)
                raise

    async def _read_driver_last_error(self) -> str | None:
        """Return ``driver.get_last_error()`` if the driver has it, else None.

        Wrapped so the error-handling path can never itself crash on a
        broken / stub driver. Called from inside ``self._lock``.
        """
        driver = self._driver
        if driver is None:
            return None
        getter = getattr(driver, "get_last_error", None)
        if getter is None:
            return None
        try:
            result = await self._io(getter)
        except Exception:
            return None
        if result is None:
            return None
        text = str(result).strip()
        return text or None

    # ---- status (side-effect-free) ----------------------------------------

    async def get_status(self) -> EquipmentStatus:
        """Produce a fresh status snapshot. MUST NOT mutate hardware state.

        The spec requires this endpoint to be safe to call every 2-3
        seconds and to always return HTTP 200 unless the process itself
        is broken. We therefore catch every per-getter failure and fold
        it into ``equipment_status: degraded`` rather than raising.

        The state lock is held only long enough to snapshot the in-memory
        bookkeeping. The instrument readback then runs outside it, on a
        worker thread, under the COM-channel lock — so a poll cannot stall a
        concurrent ``/control/*`` call, and (crucially for v1.2) a poll
        issued *during* a seal cycle answers instead of queueing behind the
        blocking ``StartCycle``.

        v1.1 ``details.claimed_by`` is attached last; the claim store has its
        own (cheap) async lock.
        """
        async with self._lock:
            driver = self._driver
            connected = self._driver_connected()
            busy = self._busy_state
            stage_state = self._stage_state
            last_error = self._last_error
            cycle_started_at = self._cycle_started_at

        if driver is None or not connected:
            readings: dict[str, Any] = {}
            readback_errors: list[str] = []
            readings_at: datetime | None = None
            stale = False
        elif busy:
            # A seal cycle owns the COM channel: reading now would queue
            # behind the blocking StartCycle and the poll would return only
            # after the cycle ended (or time out at the dashboard). Serve the
            # last observation instead, stamped with when it was taken, so
            # the reader still gets a truthful `busy` + `running` envelope.
            readings, readback_errors, readings_at = self._last_readings()
            stale = readings_at is not None
        else:
            readings, readback_errors = await self._io(
                _read_driver_metrics, driver
            )
            readings_at = datetime.now(timezone.utc)
            self._store_readings(readings, readback_errors, readings_at)
            stale = False

        status = self._compose_status(
            connected=connected,
            busy=busy,
            stage_state=stage_state,
            last_error=last_error,
            cycle_started_at=cycle_started_at,
            readings=readings,
            readback_errors=readback_errors,
            readings_at=readings_at,
            readings_stale=stale,
        )
        claimed_by = await self.claims.current()
        if claimed_by is not None:
            status.details["claimed_by"] = claimed_by.model_dump(mode="json")
        return status

    def _compose_status(
        self,
        *,
        connected: bool,
        busy: bool,
        stage_state: str,
        last_error: ErrorInfo | None,
        cycle_started_at: datetime | None,
        readings: dict[str, Any],
        readback_errors: list[str],
        readings_at: datetime | None,
        readings_stale: bool,
    ) -> EquipmentStatus:
        now = datetime.now(timezone.utc)
        uptime = time.monotonic() - self._started_at
        host = socket.gethostname()

        # ---- not connected: requires_init --------------------------------
        if not connected:
            # §2.3 invariant: requires_init ⇒ idle. Disconnected hardware
            # cannot be sealing under our control.
            self._note_activity("idle")
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
                allowed_actions=_compute_allowed_actions(
                    "requires_init",
                    "idle",
                    stage_state="unknown",
                    seal_start_blocked=True,
                ),
                activity="idle",
                activity_since=self._activity_since,
                device_time=now,
                uptime_seconds=uptime,
                last_error=last_error,
            )

        # ---- fold the readback (taken by the caller) into the envelope ----
        metrics: dict[str, MetricValue] = {}
        details: dict[str, Any] = {}
        # Metric timestamps carry the instant the values were *read*, which
        # is not `now` when a seal cycle owns the COM channel and we are
        # serving the last observation.
        read_at = readings_at or now

        actual_temp = readings.get("actual_temperature")
        if actual_temp is not None:
            metrics["actual_temperature"] = MetricValue(
                value=actual_temp, unit="C", timestamp=read_at
            )
        setpoint = readings.get("setpoint_temperature")
        if setpoint is not None:
            metrics["setpoint_temperature"] = MetricValue(
                value=setpoint, unit="C", timestamp=read_at
            )

        # Synthesized: signed delta and heater state. The ActiveX has no
        # native "ready to seal" signal so the device computes it from
        # the two raw metrics using the operator-facing tolerance. delta
        # is `actual - setpoint`, so a negative value means "still heating
        # up", positive means "above setpoint / cooling".
        actual_f = _coerce_float(actual_temp)
        setpoint_f = _coerce_float(setpoint)
        heater_temp_delta: float | None
        if actual_f is not None and setpoint_f is not None:
            heater_temp_delta = actual_f - setpoint_f
        else:
            heater_temp_delta = None
        if heater_temp_delta is not None:
            metrics["temperature_delta_c"] = MetricValue(
                value=round(heater_temp_delta, 1), unit="C", timestamp=read_at
            )
        seal_time = readings.get("sealing_time")
        if seal_time is not None:
            metrics["sealing_time"] = MetricValue(
                value=seal_time, unit="s", timestamp=read_at
            )
        cycle_count = readings.get("cycle_count")
        if cycle_count is not None:
            # `cycle_count` is the instrument's lifetime odometer, kept
            # unchanged for existing readers. `cycles_total` is the spec's
            # reserved key for the same number (§2.3.1): a seal cycle lasts
            # a few seconds, far under the dashboard's 60 s poll, so the
            # poll-to-poll delta of this counter is the *only* way a reader
            # can account for cycles it slept through. A lifetime hardware
            # counter satisfies the monotonic semantics by construction.
            metrics["cycle_count"] = MetricValue(value=cycle_count, unit="count")
            metrics["cycles_total"] = MetricValue(value=cycle_count, unit="count")

        firmware = readings.get("firmware_version")
        if firmware:
            details["firmware_version"] = firmware
        ax_version = readings.get("activex_version")
        if ax_version:
            details["activex_version"] = ax_version
        if self._connect_profile:
            details["profile"] = self._connect_profile
        com_port = readings.get("com_port")
        if com_port:
            details["com_port"] = com_port

        # ---- components --------------------------------------------------
        sealer_state = "busy" if busy else "idle"

        # Heater state is derived from the temperature delta against the
        # configured tolerance. "stable" means the plate is at setpoint
        # within +/- temperature_tolerance_c and a seal cycle would seal
        # at the requested temperature. "heating"/"cooling" mean it is
        # not yet there. "unknown" covers the case where one of the
        # readings could not be obtained.
        heater_message: str | None
        if heater_temp_delta is None:
            heater_state = "unknown"
            heater_message = None
        elif abs(heater_temp_delta) <= self._temp_tolerance_c:
            heater_state = "stable"
            heater_message = f"At setpoint ({actual_temp} C)"
        elif heater_temp_delta < 0:
            heater_state = "heating"
            heater_message = f"Heating {actual_temp} -> {setpoint} C"
        else:
            heater_state = "cooling"
            heater_message = f"Cooling {actual_temp} -> {setpoint} C"

        # Stage state is command-tracked (v1.3.0). It is snapshotted with
        # the rest of the in-memory state; on a disconnected driver we
        # cannot vouch for the carriage at all, and that path returned
        # `requires_init` above.
        stage_component_state = stage_state

        components: dict[str, ComponentStatus] = {
            "sealer": ComponentStatus(
                connected=connected,
                state=sealer_state,
            ),
            "heater": ComponentStatus(
                connected=connected,
                state=heater_state,
                message=heater_message,
                # last_event_at is intentionally None: it should be the
                # time of the last *transition* (e.g. heating -> stable),
                # not the poll timestamp. Wire that up when we have a
                # transition tracker.
            ),
            "stage": ComponentStatus(
                connected=connected, state=stage_component_state
            ),
        }

        # Tell readers what tolerance defines "stable" so a dashboard or
        # workflow can render the delta meaningfully without guessing.
        details["temperature_tolerance_c"] = self._temp_tolerance_c

        # ---- activity (v1.2 §2.3) ----------------------------------------
        # Observed from the seal-cycle state machine — `_busy_state` is true
        # for exactly the span of the blocking `StartCycle` COM call — never
        # derived from `equipment_status`. The Agilent control exposes no
        # "cycle in progress" query, so this is command-tracked, the same
        # mechanism the stage position uses. Consequence: a process restart
        # in the middle of a cycle reports `idle`; the window is the length
        # of one cycle (<= 12 s).
        activity = "running" if busy else "idle"
        # Reconcile the stored span. The transition itself is stamped by
        # start_cycle / stop_cycle / startup / shutdown; this only catches a
        # value the mutators never saw. Skipped when the snapshot has already
        # been overtaken by a concurrent transition — that mutator owns the
        # stamp and we must not overwrite it with a stale observation.
        if busy == self._busy_state:
            self._note_activity(activity)
        if cycle_started_at is not None:
            details["cycle_started_at"] = cycle_started_at.isoformat()
        if readings_stale and readings_at is not None:
            # Be explicit that the instrument values are the last observation
            # rather than a fresh read (the metric timestamps say so too).
            details["readings_as_of"] = readings_at.isoformat()

        # ---- top-level equipment_status ----------------------------------
        # Health first (§2.2), activity second (§2.3): a cycle in flight no
        # longer masks an active fault. Pre-v1.2 `busy` was tested before the
        # error/readback branches, so a fault that landed mid-cycle was
        # invisible until the cycle ended.
        if self.dry_run:
            state: str = "dry_run"
            details["dry_run"] = True
            message: str | None = (
                "[dry-run] seal cycle in progress"
                if activity == "running"
                else "Dry-run mode - no hardware connected"
            )
        elif last_error is not None and (
            (now - last_error.timestamp).total_seconds() < _RECENT_ERROR_WINDOW_S
        ):
            state = "error"
            message = last_error.message
        elif readback_errors:
            state = "degraded"
            message = "; ".join(readback_errors)
            if activity == "running":
                message += " — seal cycle continues"
        elif activity == "running":
            # Healthy + running ≡ `busy` (§2.3 invariant).
            state = "busy"
            message = "Seal cycle in progress"
        else:
            state = "ready"
            message = "Idle, ready to seal"

        # ---- allowed_actions ---------------------------------------------
        # One pure function, fed by the SAME interlock helpers the
        # /control/seal/start 412 path uses, so a workflow client trusting
        # allowed_actions verbatim cannot round-trip into a 412 the device
        # would have refused (§6.2).
        stage_blocks, _ = self.evaluate_stage_interlock(stage_component_state)
        health_blocks, _ = self.evaluate_health_interlock(last_error)
        temp_blocks, _ = self.evaluate_temperature_interlock(actual_f, setpoint_f)
        allowed_actions = _compute_allowed_actions(
            state,
            activity,
            stage_state=stage_component_state,
            seal_start_blocked=stage_blocks or health_blocks or temp_blocks,
        )

        # §6 diagnosis. An operational failure always wins; otherwise surface
        # an *active* readback fault under a stable code.
        #
        # Deliberately computed AFTER the state decision and the gates above,
        # which key off the operational `last_error` alone: folding a readback
        # warning into that variable would push it through the `error` branch
        # and through the health interlock, withholding the run for a fault
        # that does not make sealing unsafe. `degraded` + a warning is the
        # correct reading, and a warning here does not soften it (§2.3's
        # prohibition on hiding a fault).
        envelope_last_error = last_error or _readback_error_info(
            readback_errors, now
        )

        return EquipmentStatus(
            protocol_version=PROTOCOL_VERSION,
            equipment_id=self.equipment_id,
            equipment_name=self.equipment_name,
            equipment_kind=self.equipment_kind,  # type: ignore[arg-type]
            equipment_version=self.equipment_version,
            host=host,
            equipment_status=state,  # type: ignore[arg-type]
            message=message,
            allowed_actions=allowed_actions,
            activity=activity,  # type: ignore[arg-type]
            activity_since=self._activity_since,
            device_time=now,
            uptime_seconds=uptime,
            components=components,
            metrics=metrics,
            last_error=envelope_last_error,
            details=details,
        )

    # ---- helpers -----------------------------------------------------------

    async def _io(self, fn: Callable[..., Any], *args: Any) -> Any:
        """Run a blocking COM transaction on a worker thread, serialised
        through ``self._io_lock``.

        Centralises the pattern so every instrument transaction in the
        service is serialised at the COM channel: the ActiveX control is
        single-threaded, and the 32-bit surrogate serves one request at a
        time over its pipe. Lock ordering is always state -> io; nothing in
        this module acquires the state lock while holding io.
        """
        async with self._io_lock:
            return await asyncio.to_thread(fn, *args)

    def _note_activity(self, activity: str) -> None:
        """Record an observed activity value, stamping ``activity_since`` at
        the instant the value changes (§2.3: the start of the CURRENT span,
        not of the enclosing request or process)."""
        if activity != self._activity:
            self._activity = activity
            self._activity_since = datetime.now(timezone.utc)

    def _store_readings(
        self,
        readings: dict[str, Any],
        readback_errors: list[str],
        taken_at: datetime,
    ) -> None:
        self._readings = readings
        self._readback_errors = readback_errors
        self._readings_at = taken_at

    def _last_readings(self) -> tuple[dict[str, Any], list[str], datetime | None]:
        return dict(self._readings), list(self._readback_errors), self._readings_at

    def _invalidate_readings(self) -> None:
        self._readings = {}
        self._readback_errors = []
        self._readings_at = None

    def _driver_connected(self) -> bool:
        """Driver is connected if either flag is set. The real PlateLoc
        uses the private `_connected` attribute; the stub also exposes it
        for parity. Wrapped in getattr so an unexpected driver type
        cannot crash the status endpoint."""
        if self._driver is None:
            return False
        return bool(getattr(self._driver, "_connected", False))

    def clear_last_error_on_success(self) -> None:
        """Drop ``self._last_error`` after a 2xx operational response.

        Policy (single source of truth — see README "Safety interlocks"):

        * Called by every operational ``/control/*`` endpoint right
          before it returns a 2xx response (startup, shutdown,
          seal.start, seal.stop, seal.set_temperature, seal.set_time,
          stage.in, stage.out). Doing this at the API layer — not
          inside the service methods — means a refusal mid-endpoint
          (e.g. ``set_sealing_time`` succeeds then ``start_cycle``
          raises 412) does NOT clear: only an *overall* 2xx clears.
        * NOT called on 4xx / 5xx responses. A 412 from the temperature
          interlock is a refusal, not a recovery; ``last_error`` keeps
          its relevance.
        * NOT called by ``/control/claim``, ``/control/heartbeat``, or
          ``/control/release``: those are claim infrastructure, not
          operational progress, and clearing on them would hide
          ``last_error`` during a heartbeat-only retry loop.
        * NOT called by ``/status``, ``/``, or ``/health`` — those are
          read-only and must not mutate state.

        Concurrency: attribute assignment is atomic in CPython, so
        this method does NOT take ``self._lock``. A concurrent
        ``/status`` poll between the lock release and this clear sees
        the old value, which is the same staleness already inherent in
        polling.
        """
        self._last_error = None

    def _record_error(
        self, exc: Exception, method_name: str, *, detail: str | None = None
    ) -> None:
        """Capture a driver / service failure into ``self._last_error``.

        ``method_name`` is the failing service method (``"startup"``,
        ``"start_cycle"``, ``"move_stage_in"``, etc.) — used by the
        classifier as a *context* fallback when the driver text alone
        isn't decisive. The persisted ``code`` is the
        :data:`LAST_ERROR_CODES` value, not the method name.
        """
        message = str(exc)
        if detail and detail not in message:
            # The driver-reported text is what makes 0x80040201 actionable
            # ("Could not initialize - No response from PlateLoc",
            # "Stage cannot move - press is down", etc.). Append it once.
            message = f"{message} (driver: {detail})"
        code = _classify_error(method_name, exc, detail)
        self.set_last_error(code=code, message=message, severity="error")
        if detail:
            logger.exception(
                "PlateLoc error in %s (code=%s, driver: %s)",
                method_name,
                code,
                detail,
            )
        else:
            logger.exception(
                "PlateLoc error in %s (code=%s)", method_name, code
            )

    def set_last_error(
        self,
        *,
        code: str,
        message: str,
        severity: str = "error",
    ) -> None:
        """Single chokepoint for mutating ``self._last_error``.

        Enforces that ``code`` is one of :data:`LAST_ERROR_CODES` so
        callers can't accidentally introduce a free-form value that
        dashboards would then have to start string-matching on.
        Falling out of the enum is a developer error (raise immediately)
        — not a runtime condition we recover from.

        Use ``"com_other"`` as the explicit catch-all if you genuinely
        have no better classification; that is the signal to add a new
        code if the same failure mode shows up repeatedly.
        """
        if code not in LAST_ERROR_CODES:
            raise ValueError(
                f"last_error code {code!r} is not in LAST_ERROR_CODES. "
                f"Add it to the taxonomy or use 'com_other' as the "
                f"catch-all."
            )
        self._last_error = ErrorInfo(
            code=code,
            message=message,
            severity=severity,
            timestamp=datetime.now(timezone.utc),
        )


#: Instrument values folded into ``metrics``. A failed read here drives
#: ``equipment_status: degraded`` but never fails the snapshot.
_READ_LABELS = (
    ("actual_temperature", "get_actual_temperature"),
    ("setpoint_temperature", "get_sealing_temperature"),
    ("sealing_time", "get_sealing_time"),
    ("cycle_count", "get_cycle_count"),
)

#: Identity strings folded into ``details``.
_INFO_LABELS = (
    ("firmware_version", "get_firmware_version"),
    ("activex_version", "get_version"),
)


def _read_driver_metrics(driver: Any) -> tuple[dict[str, Any], list[str]]:
    """Read live driver values for the status snapshot.

    Runs on a worker thread (via :meth:`PlateLocService._io`) so the
    blocking COM transactions don't stall the event loop. Each read is
    independently try/except'd: a failed read shows up in
    ``readback_errors`` (driving ``equipment_status="degraded"``) but does
    not abort the snapshot — ``/status`` must stay a 200.
    """
    readings: dict[str, Any] = {}
    readback_errors: list[str] = []
    for label, attr in _READ_LABELS + _INFO_LABELS:
        try:
            readings[label] = getattr(driver, attr)()
        except Exception as exc:
            readback_errors.append(f"{label}: {exc}")
    com_port = getattr(driver, "com_port", None)
    if com_port:
        readings["com_port"] = com_port
    return readings, readback_errors


__all__ = [
    "LAST_ERROR_CODES",
    "PlateLocService",
    "RecentFailureNotCleared",
    "StageNotLoaded",
    "TemperatureOutOfBand",
    "_StubPlateLoc",
]
