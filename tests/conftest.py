"""Shared pytest fixtures.

The API tests run the FastAPI app with ``dry_run=True`` so no Windows /
ActiveX dependencies are required. ``conftest.py`` keeps that switch
out of every individual test.
"""

from __future__ import annotations

from collections.abc import Iterator

import pytest
from fastapi.testclient import TestClient

from agilent_plateloc.api import create_app


@pytest.fixture
def client() -> Iterator[TestClient]:
    """A `TestClient` whose lifespan auto-connects the dry-run stub.

    Each test gets a fresh app/service - no shared state across tests.
    """
    app = create_app(dry_run=True)
    with TestClient(app) as c:
        yield c
