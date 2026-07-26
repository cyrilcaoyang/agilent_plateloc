"""Lab equipment status spec types — from the shared ``sdl-lab-contract``
package.

This module used to be a verbatim vendored copy of the unified status
contract from the ac-organic-lab monorepo (``docs/STATUS_SPEC.md``). The
promised shared package now exists (``sdl-lab-contract``, versioned in
lockstep with the spec), so the types are imported and re-exported here to
keep every ``from .models import ...`` in this repo working unchanged.

Conformance: agilent_plateloc_server REST API conforms to **lab status spec
v1.2** (``activity`` / ``activity_since`` observed from the seal-cycle state
machine, the reserved ``cycles_total`` metric, and ``allowed_actions`` gated
on activity as well as on the §6 interlocks). ``PROTOCOL_VERSION`` below is
the version THIS device speaks and deliberately overrides the package's
parse-time default ("1.0" — the honest reading of a device that does not
state a version).

Kept local: :class:`ClaimRequest`, because this device validates incoming
claim bodies more strictly (field lengths, TTL bounds) than the wire
contract requires.
"""

from __future__ import annotations

from pydantic import BaseModel, Field

from sdl_lab_contract import (
    Activity,
    ClaimedBy,
    ClaimRejection,
    ClaimResponse,
    ComponentStatus,
    EquipmentKind,
    EquipmentState,
    EquipmentStatus,
    ErrorInfo,
    HealthResponse,
    MetricValue,
    ProbeResponse,
)

PROTOCOL_VERSION = "1.2"


class ClaimRequest(BaseModel):
    """Body of ``POST /control/claim`` — device-side strict validation.

    Same shape as ``sdl_lab_contract.ClaimRequest``; this device additionally
    bounds the field lengths and clamps the requested TTL.
    """

    owner: str = Field(min_length=1, max_length=120)
    session_id: str = Field(min_length=1, max_length=120)
    ttl_s: float = Field(default=30.0, ge=1.0, le=600.0)


__all__ = [
    "Activity",
    "ClaimRejection",
    "ClaimRequest",
    "ClaimResponse",
    "ClaimedBy",
    "ComponentStatus",
    "EquipmentKind",
    "EquipmentState",
    "EquipmentStatus",
    "ErrorInfo",
    "HealthResponse",
    "MetricValue",
    "PROTOCOL_VERSION",
    "ProbeResponse",
]
