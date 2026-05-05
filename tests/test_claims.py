"""STATUS_SPEC v1.1 claim-protocol conformance tests.

Covers:

* ``POST /control/claim``       - happy path, idempotent re-claim, conflict
* ``POST /control/heartbeat``   - happy path, unknown/expired token
* ``POST /control/release``     - idempotent, clears the live claim
* ``X-Claim-Token`` enforcement on every other ``/control/*`` (HTTP 423)
* ``details.claimed_by`` published on ``/status`` while a claim is held
* Advisory mode (``enforce_claims=False``) keeps publishing claim metadata
  but does not block control writes that omit the header
"""

from __future__ import annotations

from datetime import datetime, timezone

from fastapi.testclient import TestClient


# ---------------------------------------------------------------------------
# acquire / heartbeat / release happy path
# ---------------------------------------------------------------------------


def test_claim_acquire_returns_token(unclaimed_client: TestClient) -> None:
    r = unclaimed_client.post(
        "/control/claim",
        json={"owner": "alice", "session_id": "wf-1", "ttl_s": 30},
    )
    assert r.status_code == 200
    body = r.json()
    assert isinstance(body["claim_token"], str) and len(body["claim_token"]) >= 16
    assert body["heartbeat_interval_s"] >= 2.0
    # expires_at parses as RFC3339 UTC.
    expires_at = datetime.fromisoformat(body["expires_at"].replace("Z", "+00:00"))
    assert expires_at.tzinfo is not None
    assert expires_at > datetime.now(timezone.utc)


def test_claim_idempotent_same_session(unclaimed_client: TestClient) -> None:
    """A re-claim from the same session_id returns the *same* token with
    a refreshed expires_at, so an SDK that crashed mid-workflow can
    resume without coordination."""
    r1 = unclaimed_client.post(
        "/control/claim",
        json={"owner": "alice", "session_id": "wf-1", "ttl_s": 30},
    )
    token1 = r1.json()["claim_token"]
    expires1 = r1.json()["expires_at"]

    r2 = unclaimed_client.post(
        "/control/claim",
        json={"owner": "alice", "session_id": "wf-1", "ttl_s": 60},
    )
    assert r2.status_code == 200
    token2 = r2.json()["claim_token"]
    expires2 = r2.json()["expires_at"]
    assert token1 == token2
    # Refreshed expires_at must be no earlier than the first one.
    assert expires2 >= expires1


def test_claim_conflict_other_session(unclaimed_client: TestClient) -> None:
    """A second session colliding with a live claim gets HTTP 409 with
    the standard ClaimRejection body (claimed_by + retry_after_s)."""
    r1 = unclaimed_client.post(
        "/control/claim",
        json={"owner": "alice", "session_id": "wf-1", "ttl_s": 30},
    )
    assert r1.status_code == 200

    r2 = unclaimed_client.post(
        "/control/claim",
        json={"owner": "bob", "session_id": "wf-2", "ttl_s": 30},
    )
    assert r2.status_code == 409
    assert "Retry-After" in r2.headers
    # Per STATUS_SPEC v1.1 the rejection body sits at the top level so
    # SDK-side `response.json()["claimed_by"]` works without unwrapping.
    body = r2.json()
    assert body["claimed_by"]["session_id"] == "wf-1"
    assert body["claimed_by"]["owner"] == "alice"
    assert body["retry_after_s"] is not None
    assert body["retry_after_s"] >= 0.0
    assert isinstance(body["detail"], str)


def test_heartbeat_extends_claim(unclaimed_client: TestClient) -> None:
    r = unclaimed_client.post(
        "/control/claim",
        json={"owner": "alice", "session_id": "wf-1", "ttl_s": 30},
    )
    token = r.json()["claim_token"]
    expires1 = r.json()["expires_at"]

    r = unclaimed_client.post(
        "/control/heartbeat", headers={"X-Claim-Token": token}
    )
    assert r.status_code == 200
    body = r.json()
    assert body["claim_token"] == token
    # Heartbeat must move expires_at forward (or at least not backward).
    assert body["expires_at"] >= expires1


def test_heartbeat_unknown_token_returns_401(unclaimed_client: TestClient) -> None:
    r = unclaimed_client.post(
        "/control/heartbeat", headers={"X-Claim-Token": "garbage"}
    )
    assert r.status_code == 401


def test_heartbeat_after_release_returns_401(unclaimed_client: TestClient) -> None:
    r = unclaimed_client.post(
        "/control/claim",
        json={"owner": "alice", "session_id": "wf-1", "ttl_s": 30},
    )
    token = r.json()["claim_token"]
    unclaimed_client.post("/control/release", headers={"X-Claim-Token": token})

    r = unclaimed_client.post(
        "/control/heartbeat", headers={"X-Claim-Token": token}
    )
    assert r.status_code == 401


def test_release_is_idempotent(unclaimed_client: TestClient) -> None:
    """release returns 204 even when no claim is live, so an SDK that
    gives up mid-acquire (e.g. timed out before the response arrived)
    can safely call release on cleanup."""
    r = unclaimed_client.post(
        "/control/release", headers={"X-Claim-Token": "anything"}
    )
    assert r.status_code == 204


def test_release_clears_claim(unclaimed_client: TestClient) -> None:
    r = unclaimed_client.post(
        "/control/claim",
        json={"owner": "alice", "session_id": "wf-1", "ttl_s": 30},
    )
    token = r.json()["claim_token"]
    body = unclaimed_client.get("/status").json()
    assert body["details"].get("claimed_by") is not None

    r = unclaimed_client.post("/control/release", headers={"X-Claim-Token": token})
    assert r.status_code == 204

    body = unclaimed_client.get("/status").json()
    assert "claimed_by" not in body["details"]


# ---------------------------------------------------------------------------
# /status enrichment under v1.1
# ---------------------------------------------------------------------------


def test_status_publishes_claimed_by(unclaimed_client: TestClient) -> None:
    r = unclaimed_client.post(
        "/control/claim",
        json={"owner": "alice", "session_id": "wf-1", "ttl_s": 30},
    )
    assert r.status_code == 200

    body = unclaimed_client.get("/status").json()
    claimed_by = body["details"]["claimed_by"]
    assert claimed_by["session_id"] == "wf-1"
    assert claimed_by["owner"] == "alice"
    assert claimed_by["expires_at"]


def test_status_omits_claimed_by_when_unclaimed(
    unclaimed_client: TestClient,
) -> None:
    body = unclaimed_client.get("/status").json()
    assert "claimed_by" not in body["details"]


# ---------------------------------------------------------------------------
# X-Claim-Token enforcement on /control/* (strict mode)
# ---------------------------------------------------------------------------


def test_control_without_token_returns_423(unclaimed_client: TestClient) -> None:
    """Strict mode: every /control/* (except claim/heartbeat/release) is
    locked behind X-Claim-Token. Body shape matches v1.1 ClaimRejection
    (top-level ``detail``/``claimed_by``/``retry_after_s``)."""
    r = unclaimed_client.post("/control/seal/start", json={})
    assert r.status_code == 423
    body = r.json()
    assert "X-Claim-Token" in body["detail"]
    assert body["claimed_by"] is None  # nobody has the claim yet
    assert body["retry_after_s"] is None


def test_control_with_wrong_token_returns_423(unclaimed_client: TestClient) -> None:
    unclaimed_client.post(
        "/control/claim",
        json={"owner": "alice", "session_id": "wf-1", "ttl_s": 30},
    )
    r = unclaimed_client.post(
        "/control/seal/start",
        json={},
        headers={"X-Claim-Token": "definitely-not-the-real-token"},
    )
    assert r.status_code == 423
    body = r.json()
    # Device must advertise who currently holds the claim.
    assert body["claimed_by"]["session_id"] == "wf-1"
    assert body["claimed_by"]["owner"] == "alice"


def test_control_with_valid_token_succeeds(unclaimed_client: TestClient) -> None:
    r = unclaimed_client.post(
        "/control/claim",
        json={"owner": "alice", "session_id": "wf-1", "ttl_s": 30},
    )
    token = r.json()["claim_token"]

    r = unclaimed_client.post(
        "/control/seal/temperature",
        json={"temperature_c": 145},
        headers={"X-Claim-Token": token},
    )
    assert r.status_code == 200


def test_status_does_not_require_token(unclaimed_client: TestClient) -> None:
    """The read-only spec endpoints MUST stay reachable without a claim
    so the dashboard can keep polling regardless of who owns the device."""
    assert unclaimed_client.get("/").status_code == 200
    assert unclaimed_client.get("/health").status_code == 200
    assert unclaimed_client.get("/status").status_code == 200


# ---------------------------------------------------------------------------
# Advisory mode: enforce_claims=False
# ---------------------------------------------------------------------------


def test_advisory_mode_allows_unauthenticated_control(
    advisory_client: TestClient,
) -> None:
    """In advisory mode the device still publishes ``claimed_by`` but
    does not gate writes - useful for v1.0-era operator UIs that have
    not yet learned to acquire a claim."""
    r = advisory_client.post(
        "/control/seal/temperature", json={"temperature_c": 145}
    )
    assert r.status_code == 200


def test_advisory_mode_still_publishes_allowed_actions(
    advisory_client: TestClient,
) -> None:
    body = advisory_client.get("/status").json()
    assert body["protocol_version"] == "1.1"
    assert isinstance(body["allowed_actions"], list)
    assert len(body["allowed_actions"]) > 0
