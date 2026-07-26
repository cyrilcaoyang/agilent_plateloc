"""STATUS_SPEC v1.2 conformance tests (§2.3, §2.3.1, §6.2, §9 checklist).

What makes this device interesting for v1.2: the PlateLoc ActiveX control
runs in **blocking** mode, so ``StartCycle`` returns only when the physical
seal cycle has finished. The primary operation is therefore exactly the span
of that COM call — a couple of seconds, far under the dashboard's 60 s poll.
Two things follow, and both are pinned here:

* ``activity: "running"`` is only observable *while* the call is in flight,
  so these tests hold a cycle open on a worker thread and poll ``/status``
  from the main one. That also proves the poll no longer queues behind the
  cycle (it did before v1.2, when the state lock was held across the call).
* No reader polling at 60 s can see a 2 s cycle at all, which is why
  ``metrics["cycles_total"]`` (§2.3.1) is what makes plateloc utilization
  countable. It mirrors the instrument's lifetime odometer.

These tests run against the stub driver: no Windows / ActiveX dependency, no
hardware, and nothing here actuates a real sealer.
"""

from __future__ import annotations

import json
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Iterator

from fastapi.testclient import TestClient

from agilent_plateloc_server.api import create_app
from agilent_plateloc_server.service import (
    _StubPlateLoc,
    _compute_allowed_actions,
)

FIXTURES = Path(__file__).parent / "fixtures"

#: The §2.3 consistency invariants, as a table a reader may enforce.
REQUIRED_ACTIVITY: dict[str, set[str]] = {
    "busy": {"running"},
    "ready": {"idle"},
    "requires_init": {"idle"},
    "e_stop": {"idle"},
    "degraded": {"running", "idle"},
}


def _parse(stamp: str) -> datetime:
    """Parse a wire timestamp. Pydantic serialises datetimes with ``Z``;
    values this repo writes into ``details`` use ``isoformat()`` (``+00:00``).
    Both are ISO-8601 UTC, so compare the parsed instants."""
    return datetime.fromisoformat(stamp.replace("Z", "+00:00"))


class _HeldCycleStub(_StubPlateLoc):
    """Stub whose ``start_cycle`` blocks until the test releases it.

    Stands in for the real control's blocking ``StartCycle``: the stub's
    own version returns instantly, which is fine for state-machine tests but
    leaves no window in which to observe ``activity: "running"``.
    """

    def __init__(self) -> None:
        super().__init__()
        self.entered = threading.Event()
        self.release = threading.Event()

    def start_cycle(self) -> int:
        self.entered.set()
        if not self.release.wait(timeout=10.0):  # pragma: no cover - test bug
            raise TimeoutError("test never released the held seal cycle")
        return super().start_cycle()


def _client(
    driver_factory: type[_StubPlateLoc] = _StubPlateLoc,
    *,
    home_stage: bool = True,
) -> tuple[TestClient, _StubPlateLoc]:
    """Start a service on the given stub with ``dry_run=False`` (so the real
    operational state machine runs) and claims advisory (the claim protocol
    is covered by ``test_claims.py``). Returns ``(client, driver)``.
    """
    app = create_app(dry_run=False, enforce_claims=False)
    app.state.service._driver_factory = driver_factory
    client = TestClient(app)
    client.__enter__()  # lifespan auto-connects the stub
    if home_stage:
        client.post("/control/stage/in")
    return client, app.state.service._driver


@contextmanager
def _cycle_in_flight(
    client: TestClient, driver: _HeldCycleStub, **body: object
) -> Iterator[None]:
    """Hold one seal cycle open: POST /control/seal/start on a worker thread,
    wait until the driver is inside the blocking call, then yield. On exit the
    cycle is released and the POST is asserted to have succeeded."""
    payload = {"temperature_c": 170, "seconds": 3.0, **body}
    with ThreadPoolExecutor(max_workers=1) as pool:
        future = pool.submit(client.post, "/control/seal/start", json=payload)
        assert driver.entered.wait(timeout=5.0), "cycle never started"
        try:
            yield
        finally:
            driver.release.set()
            response = future.result(timeout=10.0)
    assert response.status_code == 200, response.text


# ---------------------------------------------------------------------------
# §2.3 — activity is observed, and the invariants hold
# ---------------------------------------------------------------------------


def test_busy_and_running_while_the_cycle_is_in_flight() -> None:
    client, driver = _client(_HeldCycleStub)
    try:
        before = client.get("/status").json()
        assert before["equipment_status"] == "ready"
        assert before["activity"] == "idle"  # ready ⇒ idle

        with _cycle_in_flight(client, driver):
            body = client.get("/status").json()
            # busy ≡ healthy + running (§2.3 invariant)
            assert body["equipment_status"] == "busy"
            assert body["activity"] == "running"
            assert body["components"]["sealer"]["state"] == "busy"
            assert body["activity_since"] is not None
            assert body["activity_since"] > before["activity_since"]
            # No second concurrent run, no carriage move; abort stays
            # reachable (§2.3).
            assert set(body["allowed_actions"]) == {"shutdown", "seal.stop"}

        after = client.get("/status").json()
        assert after["equipment_status"] == "ready"
        assert after["activity"] == "idle"
        assert "seal.start" in after["allowed_actions"]
    finally:
        client.__exit__(None, None, None)


def test_status_answers_during_a_cycle_instead_of_queueing_behind_it() -> None:
    """The poll must not wait for the cycle: before v1.2 ``/status`` took the
    same lock ``start_cycle`` held across the blocking COM call, so a reader
    could only ever see the device before or after a cycle — never running."""
    client, driver = _client(_HeldCycleStub)
    try:
        client.get("/status")  # prime the readback cache
        with _cycle_in_flight(client, driver):
            started = time.perf_counter()
            body = client.get("/status").json()
            elapsed = time.perf_counter() - started
            assert elapsed < 1.0, f"/status blocked for {elapsed:.2f}s"
            assert body["activity"] == "running"
            # Instrument values are the last observation, not a fresh read
            # (the COM channel belongs to the cycle) — and the envelope says
            # so, both in details and on every metric timestamp.
            assert "readings_as_of" in body["details"]
            assert _parse(body["metrics"]["actual_temperature"]["timestamp"]) == _parse(
                body["details"]["readings_as_of"]
            )
            assert _parse(body["details"]["readings_as_of"]) < _parse(
                body["device_time"]
            )
    finally:
        client.__exit__(None, None, None)


def test_requires_init_implies_idle() -> None:
    client, _ = _client(home_stage=False)
    try:
        client.post("/control/shutdown")
        body = client.get("/status").json()
        assert body["equipment_status"] == "requires_init"
        assert body["activity"] == "idle"
        assert body["activity_since"] is not None
        assert body["allowed_actions"] == ["startup"]
    finally:
        client.__exit__(None, None, None)


def test_activity_since_stamps_the_span_not_the_poll() -> None:
    """``activity_since`` is the instant the value last *changed*; repeated
    polls of an unchanged activity must not move it."""
    client, driver = _client(_HeldCycleStub)
    try:
        first = client.get("/status").json()["activity_since"]
        assert client.get("/status").json()["activity_since"] == first

        with _cycle_in_flight(client, driver):
            running_since = client.get("/status").json()["activity_since"]
            assert running_since > first
            assert client.get("/status").json()["activity_since"] == running_since

        idle_again = client.get("/status").json()["activity_since"]
        assert idle_again > running_since
    finally:
        client.__exit__(None, None, None)


def test_activity_is_not_derived_from_equipment_status() -> None:
    """§2.3 forbids computing one from the other. The observable proof: a
    fault that lands mid-cycle changes the health axis and leaves activity
    alone — pre-v1.2 the ``busy`` branch was tested first and hid it."""
    client, driver = _client(_HeldCycleStub)
    try:
        driver.get_sealing_time = _boom(  # type: ignore[method-assign]
            "GetSealingTime timed out"
        )
        client.get("/status")  # observe the fault while idle
        with _cycle_in_flight(client, driver):
            body = client.get("/status").json()
            assert body["equipment_status"] == "degraded"
            assert body["activity"] == "running"
            assert "seal cycle continues" in body["message"]
            # The activity gate wins over the health gate for the abort
            # class: a fault must never take seal.stop away mid-cycle.
            assert set(body["allowed_actions"]) == {"shutdown", "seal.stop"}
    finally:
        client.__exit__(None, None, None)


def test_degraded_keeps_the_capability_the_fault_does_not_touch() -> None:
    """§2.2's "safe, useful subset": a failing *seal-time* readback is a real
    fault (→ ``degraded``) but does not make sealing unsafe, so the run stays
    offered. Through v1.3.2 ``degraded`` collapsed to ``["shutdown"]`` while
    ``/control/seal/start`` still returned 200 — a §6.2 violation in the
    withholding direction.
    """
    client, driver = _client()
    try:
        driver.get_sealing_time = _boom(  # type: ignore[method-assign]
            "GetSealingTime timed out"
        )
        body = client.get("/status").json()
        assert body["equipment_status"] == "degraded"
        assert body["activity"] == "idle"  # degraded ⇒ running OR idle
        assert "GetSealingTime timed out" in body["message"]
        # The fault is also reported under a stable code, so a dashboard
        # branches on `last_error.code` instead of regexing `message`
        # (best-practice #6). Severity `warning`: `degraded` already carries
        # the safety consequence.
        assert body["last_error"]["code"] == "com_timeout"
        assert body["last_error"]["severity"] == "warning"
        # The heater is readable, so the temperature interlock passes and the
        # run stays available — and the endpoint agrees.
        assert "seal.start" in body["allowed_actions"]
        assert client.post("/control/seal/start", json={"seconds": 3.0}).status_code == 200
    finally:
        client.__exit__(None, None, None)


def test_unreadable_temperature_is_degraded_and_fails_closed() -> None:
    """The readback that *is* load-bearing: with the heater unreadable the
    temperature interlock fails closed, so the run disappears from both
    surfaces at once."""
    client, driver = _client()
    try:
        driver.get_actual_temperature = lambda: None  # type: ignore[method-assign]
        body = client.get("/status").json()
        assert body["components"]["heater"]["state"] == "unknown"
        assert "seal.start" not in body["allowed_actions"]
        r = client.post("/control/seal/start", json={"seconds": 3.0})
        assert r.status_code == 412
        assert "Cannot verify temperature" in r.json()["detail"]
    finally:
        client.__exit__(None, None, None)


# ---------------------------------------------------------------------------
# §2.3.1 — cycles_total is what makes a sub-poll-interval cycle countable
# ---------------------------------------------------------------------------


def test_cycles_total_mirrors_the_instrument_odometer() -> None:
    client, _ = _client()
    try:
        body = client.get("/status").json()
        before = body["metrics"]["cycles_total"]["value"]
        assert body["metrics"]["cycles_total"]["unit"] == "count"
        # Same number as the pre-existing `cycle_count`, which is kept
        # unchanged for existing readers.
        assert body["metrics"]["cycle_count"]["value"] == before

        # Two cycles, each shorter than any poll interval: a reader that
        # slept through both still recovers the count from the delta.
        for _ in range(2):
            r = client.post(
                "/control/seal/start", json={"temperature_c": 170, "seconds": 3.0}
            )
            assert r.status_code == 200, r.text
            client.post("/control/stage/in")  # re-home for the next cycle

        body = client.get("/status").json()
        assert body["metrics"]["cycles_total"]["value"] == before + 2
        assert body["metrics"]["cycle_count"]["value"] == before + 2

        # Stopping when nothing is running is not a completed cycle.
        assert client.post("/control/seal/stop").status_code == 200
        body = client.get("/status").json()
        assert body["metrics"]["cycles_total"]["value"] == before + 2
    finally:
        client.__exit__(None, None, None)


def test_cycles_total_survives_a_failed_cycle_without_counting_it() -> None:
    client, driver = _client()
    try:
        before = client.get("/status").json()["metrics"]["cycles_total"]["value"]
        driver.start_cycle = _boom("Low Air Pressure Error")  # type: ignore[method-assign]
        assert (
            client.post(
                "/control/seal/start", json={"temperature_c": 170, "seconds": 3.0}
            ).status_code
            == 500
        )
        body = client.get("/status").json()
        assert body["metrics"]["cycles_total"]["value"] == before
        # The failure ends the activity span — a fault is not a run (§2.3).
        assert body["activity"] == "idle"
        assert body["equipment_status"] == "error"
        assert body["last_error"]["code"] == "low_air_pressure"
    finally:
        client.__exit__(None, None, None)


# ---------------------------------------------------------------------------
# §6.2 — allowed_actions and the /control/* refusals cannot disagree
# ---------------------------------------------------------------------------


def test_mid_cycle_refusals_mirror_the_advertised_list() -> None:
    """What ``/status`` withholds while a cycle runs, ``/control/*`` refuses.

    The refusals are 409, not 412: "a cycle is already running" is a
    device-state conflict, not an unmet precondition (§6.1).
    """
    client, driver = _client(_HeldCycleStub)
    try:
        with _cycle_in_flight(client, driver):
            advertised = set(client.get("/status").json()["allowed_actions"])
            assert advertised == {"shutdown", "seal.stop"}
            for action, request in (
                ("seal.start", lambda: client.post(
                    "/control/seal/start", json={"seconds": 3.0}
                )),
                ("stage.in", lambda: client.post("/control/stage/in")),
                ("stage.out", lambda: client.post("/control/stage/out")),
                ("seal.set_temperature", lambda: client.post(
                    "/control/seal/temperature", json={"temperature_c": 170}
                )),
                ("seal.set_time", lambda: client.post(
                    "/control/seal/time", json={"seconds": 2.0}
                )),
            ):
                assert action not in advertised
                response = request()
                assert response.status_code == 409, (
                    f"{action}: expected 409 while a cycle is in flight, "
                    f"got {response.status_code}"
                )
                assert "seal cycle is in progress" in response.json()["detail"]
    finally:
        client.__exit__(None, None, None)


def test_health_interlock_withholds_the_run_but_not_the_recovery() -> None:
    """After an operational failure the device must stop offering a new run
    (§2.2) *and* keep offering the actions an operator needs to recover — the
    2026-07-15 bench failure left a plate in a hot chamber while the device
    advertised nothing but ``shutdown``.

    Also the §6.4 half: the first 2xx from an operational endpoint clears the
    error, and the run comes back on both surfaces at once.
    """
    client, driver = _client()
    try:
        driver.set_sealing_time = _boom(  # type: ignore[method-assign]
            "Hot Plate Vacuum Error"
        )
        assert client.post("/control/seal/time", json={"seconds": 2.0}).status_code == 500

        body = client.get("/status").json()
        assert body["equipment_status"] == "error"
        assert body["last_error"]["code"] == "vacuum_error"
        assert "seal.start" not in body["allowed_actions"]
        # Recovery + diagnostics stay reachable.
        assert {"stage.out", "shutdown"} <= set(body["allowed_actions"])

        # The 412 mirrors the omission, with its own body shape (§6.1) and a
        # Retry-After: the window expires on its own.
        r = client.post("/control/seal/start", json={"temperature_c": 170})
        assert r.status_code == 412, r.text
        refusal = r.json()
        assert refusal["detail"] == "Recent operational failure not cleared"
        assert refusal["last_error_code"] == "vacuum_error"
        assert refusal["retry_after_s"] > 0
        assert r.headers["Retry-After"] == str(int(refusal["retry_after_s"]))

        # §6.4: the first successful operational action clears it, and both
        # surfaces recover together.
        assert client.post("/control/stage/in").status_code == 200
        body = client.get("/status").json()
        assert body["last_error"] is None
        assert body["equipment_status"] == "ready"
        assert "seal.start" in body["allowed_actions"]
        assert (
            client.post("/control/seal/start", json={"temperature_c": 170}).status_code
            == 200
        )
    finally:
        client.__exit__(None, None, None)


def test_a_failed_cycle_leaves_no_phantom_run() -> None:
    """A cycle that dies inside the COM call must clear the activity span —
    otherwise the device would advertise only the abort class forever."""
    client, driver = _client()
    try:
        driver.start_cycle = _boom("Simulated mid-cycle COM fault")  # type: ignore[method-assign]
        client.post("/control/seal/start", json={"temperature_c": 170, "seconds": 3.0})
        body = client.get("/status").json()
        assert body["activity"] == "idle"
        # Stage is pessimized to unknown: the carriage moved mid-failure.
        assert body["components"]["stage"]["state"] == "unknown"
        # A subsequent stage move is honored again (not stuck in "running").
        assert client.post("/control/stage/out").status_code == 200
    finally:
        client.__exit__(None, None, None)


def test_compute_allowed_actions_gates(  # noqa: D103 - table-driven
) -> None:
    # requires_init / e_stop / unknown
    assert _compute_allowed_actions(
        "requires_init", "idle", stage_state="unknown", seal_start_blocked=True
    ) == ["startup"]
    for state in ("e_stop", "unknown"):
        assert (
            _compute_allowed_actions(
                state, "idle", stage_state="in", seal_start_blocked=False
            )
            == []
        )

    # activity wins over health for the abort class, in every health state.
    for state in ("ready", "busy", "degraded", "error", "dry_run"):
        assert _compute_allowed_actions(
            state, "running", stage_state="in", seal_start_blocked=False
        ) == ["shutdown", "seal.stop"]

    # error / degraded keep the recovery + diagnostic actions (§2.2); only
    # the *run* is withheld, and that decision arrives via the interlocks.
    for state in ("error", "degraded"):
        recovery = _compute_allowed_actions(
            state, "idle", stage_state="unknown", seal_start_blocked=True
        )
        assert "seal.start" not in recovery
        assert {"stage.in", "stage.out", "shutdown"} <= set(recovery)

    # ready/idle: seal.start iff no interlock blocks; no-op stage direction
    # is dedup'd.
    allowed = _compute_allowed_actions(
        "ready", "idle", stage_state="in", seal_start_blocked=False
    )
    assert "seal.start" in allowed and "stage.in" not in allowed
    assert "stage.out" in allowed and "seal.stop" not in allowed
    blocked = _compute_allowed_actions(
        "ready", "idle", stage_state="out", seal_start_blocked=True
    )
    assert "seal.start" not in blocked
    assert "stage.in" in blocked and "stage.out" not in blocked


# ---------------------------------------------------------------------------
# §9 v1.2 checklist — snapshot fixtures
# ---------------------------------------------------------------------------


def _boom(message: str):
    """Return a callable that raises ``OSError`` — the shape a real ActiveX
    COM fault takes (``pywintypes.com_error``), not ``RuntimeError``."""

    def _raise(*_a: object, **_kw: object) -> None:
        raise OSError(message)

    return _raise


def test_save_v12_status_fixtures(scrub) -> None:
    """Write the two v1.2 activity snapshots (§9 asks for healthy+running,
    healthy+idle, and a degraded shape; healthy+idle is
    ``status_ready.json``, written by ``test_api.py``).

      - status_busy.json             - busy + running, mid-cycle
      - status_degraded_running.json - degraded + running: a heater readback
                                       fault does not hide the cycle, and the
                                       cycle does not hide the fault
    """
    FIXTURES.mkdir(exist_ok=True)

    client, driver = _client(_HeldCycleStub)
    try:
        client.get("/status")  # prime the readback cache
        with _cycle_in_flight(client, driver):
            body = client.get("/status").json()
            assert body["equipment_status"] == "busy"
            assert body["activity"] == "running"
            (FIXTURES / "status_busy.json").write_text(
                json.dumps(scrub(body), indent=2, sort_keys=True) + "\n"
            )
    finally:
        client.__exit__(None, None, None)

    client, driver = _client(_HeldCycleStub)
    try:
        driver.get_sealing_time = _boom(  # type: ignore[method-assign]
            "GetSealingTime timed out"
        )
        client.get("/status")  # observe the fault while idle
        with _cycle_in_flight(client, driver):
            body = client.get("/status").json()
            assert body["equipment_status"] == "degraded"
            assert body["activity"] == "running"
            (FIXTURES / "status_degraded_running.json").write_text(
                json.dumps(scrub(body), indent=2, sort_keys=True) + "\n"
            )
    finally:
        client.__exit__(None, None, None)


def test_fixture_snapshots_validate_against_the_contract() -> None:
    """Every checked-in fixture parses under the shared contract package and
    respects the §2.3 consistency invariants."""
    from sdl_lab_contract import EquipmentStatus

    fixtures = sorted(FIXTURES.glob("status_*.json"))
    assert len(fixtures) >= 3  # ready+idle, busy+running, degraded+running
    seen: set[tuple[str, str]] = set()
    for path in fixtures:
        status = EquipmentStatus.model_validate(json.loads(path.read_text()))
        assert status.protocol_version == "1.2", path.name
        assert status.equipment_version, path.name
        required = REQUIRED_ACTIVITY.get(status.equipment_status)
        if required is not None:
            assert status.activity in required, path.name
        seen.add((status.equipment_status, status.activity))

    assert ("ready", "idle") in seen
    assert ("busy", "running") in seen
    assert ("degraded", "running") in seen
