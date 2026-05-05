"""Conformance tests for the lab equipment status spec v1.1.

These tests run with the dry-run stub driver so they require no Windows
/ ActiveX dependencies and can be executed in CI on any platform.

The default ``client`` fixture (see ``conftest.py``) is built with
``enforce_claims=True`` and pre-acquires a claim, so v1.0-era control
tests keep working unchanged. v1.1-specific surface (claim protocol,
``allowed_actions``, ``details.claimed_by``, 423 enforcement) is
covered here and in ``test_claims.py``.
"""

from __future__ import annotations

import json
from pathlib import Path

from fastapi.testclient import TestClient

from agilent_plateloc.models import PROTOCOL_VERSION

FIXTURES = Path(__file__).parent / "fixtures"


# ---------------------------------------------------------------------------
# Spec endpoints
# ---------------------------------------------------------------------------


def test_probe(client: TestClient) -> None:
    r = client.get("/")
    assert r.status_code == 200
    body = r.json()
    assert body["equipment_id"] == "plateloc"
    assert body["equipment_name"] == "Agilent PlateLoc"
    assert body["protocol_version"] == PROTOCOL_VERSION
    assert body["protocol_version"] == "1.1"


def test_health(client: TestClient) -> None:
    r = client.get("/health")
    assert r.status_code == 200
    assert r.json() == {"status": "healthy"}


def test_openapi_doc(client: TestClient) -> None:
    """FastAPI auto-publishes /openapi.json - the spec requires it.

    We assert that the v1.0 + v1.1 schemas are in the doc so an
    ``openapi-typescript`` consumer (e.g. the dashboard frontend) can
    pull types straight from this device.
    """
    r = client.get("/openapi.json")
    assert r.status_code == 200
    schemas = r.json()["components"]["schemas"]
    for required in [
        # v1.0 envelope
        "EquipmentStatus",
        "ProbeResponse",
        "HealthResponse",
        "ComponentStatus",
        "MetricValue",
        "ErrorInfo",
        # v1.1 claim protocol
        "ClaimRequest",
        "ClaimResponse",
        "ClaimRejection",
        "ClaimedBy",
    ]:
        assert required in schemas, f"OpenAPI doc is missing {required}"


def test_status_envelope(client: TestClient) -> None:
    """Spec-required fields exist and have the correct types/shape."""
    r = client.get("/status")
    assert r.status_code == 200
    body = r.json()

    assert body["protocol_version"] == PROTOCOL_VERSION
    assert body["equipment_id"] == "plateloc"
    assert body["equipment_kind"] == "plate_sealer"
    assert body["equipment_status"] == "dry_run"
    assert isinstance(body["device_time"], str)
    assert isinstance(body["uptime_seconds"], (int, float))

    # v1.1: allowed_actions is a top-level list of skill names.
    assert isinstance(body["allowed_actions"], list)
    assert "seal.start" in body["allowed_actions"]
    assert "shutdown" in body["allowed_actions"]

    # Metrics are populated from the stub driver.
    metrics = body["metrics"]
    assert metrics["actual_temperature"]["unit"] == "C"
    assert metrics["setpoint_temperature"]["unit"] == "C"
    assert metrics["sealing_time"]["unit"] == "s"
    assert metrics["cycle_count"]["unit"] == "count"

    # Components.
    assert "sealer" in body["components"]
    assert "stage" in body["components"]


def test_status_is_side_effect_free(client: TestClient) -> None:
    """Spec rule #1: GET /status MUST be side-effect-free.

    Polling repeatedly must not increment the cycle counter or otherwise
    mutate state.
    """
    r1 = client.get("/status")
    cc1 = r1.json()["metrics"]["cycle_count"]["value"]
    for _ in range(10):
        client.get("/status")
    r2 = client.get("/status")
    cc2 = r2.json()["metrics"]["cycle_count"]["value"]
    assert cc1 == cc2 == 0


def test_status_always_200_when_disconnected(unclaimed_client: TestClient) -> None:
    """Spec rule #2: /status returns 200 even if hardware isn't ready.

    We force a disconnect by claiming + calling /control/shutdown, then
    verify the response is HTTP 200 with `equipment_status: requires_init`.
    """
    r = unclaimed_client.post(
        "/control/claim",
        json={"owner": "pytest", "session_id": "shutdown-test", "ttl_s": 30},
    )
    token = r.json()["claim_token"]
    unclaimed_client.post("/control/shutdown", headers={"X-Claim-Token": token})
    r = unclaimed_client.get("/status")
    assert r.status_code == 200
    body = r.json()
    assert body["equipment_status"] == "requires_init"
    assert "startup" in body["required_actions"]
    # In requires_init the only action the device will honour is startup.
    assert body["allowed_actions"] == ["startup"]


# ---------------------------------------------------------------------------
# v1.1 allowed_actions semantics
# ---------------------------------------------------------------------------


def test_allowed_actions_changes_with_state(client: TestClient) -> None:
    """allowed_actions must reflect current equipment_status.

    Walks the dry-run state machine: dry_run starts with the full set
    (because dry_run is by definition able to honour everything), then
    after explicit shutdown we switch to requires_init -> startup-only.
    """
    body = client.get("/status").json()
    assert body["equipment_status"] == "dry_run"
    assert "seal.start" in body["allowed_actions"]
    assert "stage.in" in body["allowed_actions"]

    client.post("/control/shutdown")
    body = client.get("/status").json()
    assert body["equipment_status"] == "requires_init"
    assert body["allowed_actions"] == ["startup"]


def test_allowed_actions_ready_state() -> None:
    """ready (real driver, not dry_run) exposes the full operating set
    minus seal.stop (which only makes sense while busy)."""
    from agilent_plateloc.api import create_app
    from agilent_plateloc.service import _StubPlateLoc

    app = create_app(dry_run=False, enforce_claims=True)
    app.state.service._driver_factory = _StubPlateLoc
    with TestClient(app) as alt:
        r = alt.post(
            "/control/claim",
            json={"owner": "pytest", "session_id": "ready-test", "ttl_s": 60},
        )
        token = r.json()["claim_token"]
        alt.headers["X-Claim-Token"] = token

        alt.post("/control/startup", json={})
        body = alt.get("/status").json()
        assert body["equipment_status"] == "ready"
        actions = set(body["allowed_actions"])
        assert {"seal.start", "stage.in", "stage.out", "shutdown"} <= actions
        assert "seal.stop" not in actions  # nothing to stop yet


def test_allowed_actions_busy_state() -> None:
    """busy advertises seal.stop and shutdown (and nothing that would
    conflict with an in-flight cycle)."""
    from agilent_plateloc.api import create_app
    from agilent_plateloc.service import _StubPlateLoc

    app = create_app(dry_run=False, enforce_claims=True)
    app.state.service._driver_factory = _StubPlateLoc
    with TestClient(app) as alt:
        r = alt.post(
            "/control/claim",
            json={"owner": "pytest", "session_id": "busy-test", "ttl_s": 60},
        )
        alt.headers["X-Claim-Token"] = r.json()["claim_token"]

        alt.post("/control/startup", json={})
        alt.post("/control/seal/start", json={"temperature_c": 170, "seconds": 3.0})
        body = alt.get("/status").json()
        assert body["equipment_status"] == "busy"
        actions = set(body["allowed_actions"])
        assert "seal.stop" in actions
        assert "shutdown" in actions
        assert "seal.start" not in actions  # already running


# ---------------------------------------------------------------------------
# Control endpoints (existing v1.0 behaviour, now under a held claim)
# ---------------------------------------------------------------------------


def test_set_temperature_validation(client: TestClient) -> None:
    """Out-of-range values are rejected before they reach the driver."""
    r = client.post("/control/seal/temperature", json={"temperature_c": 500})
    assert r.status_code == 422


def test_seal_cycle_round_trip(client: TestClient) -> None:
    """Start a cycle, then stop it, and confirm the cycle counter
    incremented exactly once."""
    before = client.get("/status").json()["metrics"]["cycle_count"]["value"]

    r = client.post(
        "/control/seal/start", json={"temperature_c": 170, "seconds": 3.0}
    )
    assert r.status_code == 200

    r = client.post("/control/seal/stop")
    assert r.status_code == 200

    after = client.get("/status").json()["metrics"]["cycle_count"]["value"]
    assert after == before + 1


def test_temperature_setpoint_persists(client: TestClient) -> None:
    """A temperature set via /control/seal/temperature is visible in
    the next /status response."""
    r = client.post("/control/seal/temperature", json={"temperature_c": 145})
    assert r.status_code == 200
    body = client.get("/status").json()
    assert body["metrics"]["setpoint_temperature"]["value"] == 145


def test_shutdown_then_control_returns_409(client: TestClient) -> None:
    """Spec-friendly behaviour: control endpoints fail with 409 (not 500)
    when the driver isn't connected, so the operator UI can render a
    clear "click Connect first" message.

    Note: the claim is acquired by the fixture, so the 423 path is *not*
    hit here; this test exists to assert the post-shutdown 409 path,
    which is independent of v1.1 enforcement.
    """
    client.post("/control/shutdown")
    r = client.post("/control/seal/start", json={})
    assert r.status_code == 409


# ---------------------------------------------------------------------------
# Snapshot fixtures (saved for regression review)
# ---------------------------------------------------------------------------


def _scrub_for_diff(body: dict) -> dict:
    """Replace runtime-volatile fields with stable placeholders so the
    saved fixtures only diff when the schema or value semantics change."""
    body["device_time"] = "2026-04-29T22:50:01Z"
    body["uptime_seconds"] = 0.0
    body["host"] = "plateloc-pc"
    for metric in body.get("metrics", {}).values():
        if metric.get("timestamp"):
            metric["timestamp"] = "2026-04-29T22:50:01Z"
    # Claim expiry is wall-clock; scrub the same way as device_time.
    if isinstance(body.get("details"), dict) and "claimed_by" in body["details"]:
        body["details"]["claimed_by"]["expires_at"] = "2026-04-29T22:51:01Z"
    return body


def test_save_status_fixtures(unclaimed_client: TestClient) -> None:
    """Re-generate ``tests/fixtures/status_*.json``.

    Fixtures are checked into git so reviewers can eyeball schema
    changes. After intentional schema changes, re-run pytest and commit
    the diffs as part of the PR.

    Coverage:
      - status_requires_init.json   - hardware not connected (spec example)
      - status_ready.json           - connected & idle (uses stub driver)
      - status_ready_claimed.json   - same, but with details.claimed_by
      - status_busy.json            - cycle in progress (uses stub driver)
      - status_dry_run.json         - dry-run mode advertised in /status
    """
    from agilent_plateloc.api import create_app
    from agilent_plateloc.service import _StubPlateLoc

    FIXTURES.mkdir(exist_ok=True)

    # dry_run snapshot (no claim active).
    body = unclaimed_client.get("/status").json()
    (FIXTURES / "status_dry_run.json").write_text(
        json.dumps(_scrub_for_diff(body), indent=2, sort_keys=True) + "\n"
    )

    # ready/busy: spin up a fresh service with the stub injected via
    # driver_factory but `dry_run=False`, so equipment_status reflects
    # the real operational state machine.
    app = create_app(dry_run=False, enforce_claims=True)
    app.state.service._driver_factory = _StubPlateLoc
    with TestClient(app) as alt:
        # Acquire the claim under a stable session_id so the fixture is
        # reproducible (apart from expires_at, which the scrubber pins).
        r = alt.post(
            "/control/claim",
            json={"owner": "fixture", "session_id": "fixture-session", "ttl_s": 60},
        )
        token = r.json()["claim_token"]
        alt.headers["X-Claim-Token"] = token

        alt.post("/control/startup", json={})
        body = alt.get("/status").json()
        assert body["equipment_status"] == "ready"
        # Snapshot WITH the claim metadata so reviewers see the v1.1 shape.
        (FIXTURES / "status_ready_claimed.json").write_text(
            json.dumps(_scrub_for_diff(body), indent=2, sort_keys=True) + "\n"
        )

        # Snapshot WITHOUT claim metadata (back-compat with v1.0 readers).
        # Release the claim, re-poll, snapshot.
        alt.post("/control/release", headers={"X-Claim-Token": token})
        del alt.headers["X-Claim-Token"]
        body = alt.get("/status").json()
        assert "claimed_by" not in body["details"]
        (FIXTURES / "status_ready.json").write_text(
            json.dumps(_scrub_for_diff(body), indent=2, sort_keys=True) + "\n"
        )

        # Re-acquire for the busy snapshot.
        r = alt.post(
            "/control/claim",
            json={
                "owner": "fixture",
                "session_id": "fixture-session",
                "ttl_s": 60,
            },
        )
        alt.headers["X-Claim-Token"] = r.json()["claim_token"]

        alt.post(
            "/control/seal/start", json={"temperature_c": 170, "seconds": 3.0}
        )
        body = alt.get("/status").json()
        assert body["equipment_status"] == "busy"
        (FIXTURES / "status_busy.json").write_text(
            json.dumps(_scrub_for_diff(body), indent=2, sort_keys=True) + "\n"
        )

    # requires_init: shut the dry-run driver down explicitly.
    r = unclaimed_client.post(
        "/control/claim",
        json={"owner": "fixture", "session_id": "fixture-shutdown", "ttl_s": 60},
    )
    unclaimed_client.headers["X-Claim-Token"] = r.json()["claim_token"]
    unclaimed_client.post("/control/shutdown")
    unclaimed_client.post(
        "/control/release",
        headers={"X-Claim-Token": unclaimed_client.headers["X-Claim-Token"]},
    )
    del unclaimed_client.headers["X-Claim-Token"]
    body = unclaimed_client.get("/status").json()
    assert body["equipment_status"] == "requires_init"
    (FIXTURES / "status_requires_init.json").write_text(
        json.dumps(_scrub_for_diff(body), indent=2, sort_keys=True) + "\n"
    )
