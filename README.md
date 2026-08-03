# Agilent PlateLoc Thermal Microplate Sealer — Python Driver + REST API

Python driver and REST API service for the **Agilent PlateLoc Thermal Microplate Sealer**, communicating through the VWorks ActiveX COM control over a serial (COM) port.

> **API conformance:** This repo conforms to **lab status spec v1.2** (see `docs/STATUS_SPEC.md` in the [`ac-organic-lab`](https://github.com/cyrilcaoyang/ac-organic-lab) monorepo; the contract types are imported from the shared [`sdl-lab-contract`](https://github.com/AccelerationConsortium/sdl-lab-contract) package). The dashboard auto-discovers this device by polling its `/status` endpoint; the SDK acquires a short-lived claim via `POST /control/claim` before issuing other `/control/*` writes.
>
> **Primary operation (v1.2 `activity`):** a **seal cycle**. `/status` reports `activity: "running"` for exactly as long as one cycle is executing and `"idle"` otherwise — see [Activity and cycle accounting](#activity-and-cycle-accounting-v12).

## Prerequisites

- **Windows** (ActiveX is Windows-only)
- **Python 3.10+**
- **Agilent VWorks ActiveX Controls** installed (from the Agilent software CD/UFD)
- **32-bit Python** installed alongside your main Python (the ActiveX DLL is 32-bit — see [32-bit note](#32-bit-python-requirement) below)
- PlateLoc connected via **RS-232 serial** (e.g. COM14 — set in `config.toml`)

## Installation

[uv](https://docs.astral.sh/uv/) is the canonical environment manager for this repo and for the rest of the [`ac-organic-lab`](https://github.com/AccelerationConsortium/ac-organic-lab) stack. It provides reproducible installs (every dependency is pinned in `uv.lock`), is significantly faster than pip, and integrates cleanly with the Windows Service supervisor used in production (NSSM — see [Production deployment](#production-deployment) below).

```powershell
# Install uv (one-time per PC)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

# Clone / navigate to the project
cd path\to\agilent-plateloc-server

# Copy the example config and edit for your setup
copy config.example.toml config.toml
# Edit config.toml — set com_port, profile name, port, enforce_claims, etc.

# Sync runtime + dev dependencies (creates .venv automatically)
uv sync --extra dev

# Run tests
uv run pytest

# Run the service in the foreground (Ctrl-C to stop)
uv run --extra api agilent-plateloc-serve
```

`uv sync --extra dev` installs everything needed to run the test suite (`pytest`, `httpx`, etc.) plus the FastAPI runtime. For a runtime-only install (e.g., on the lab PC), use `uv sync --extra api` — see [Production deployment](#production-deployment).

> **Already on conda?** If your team's standard is conda and you'd rather not introduce a second tool, see [Appendix: Alternative install via conda](#appendix-alternative-install-via-conda) at the bottom of this README. Functionally equivalent; the rest of the lab still runs uv.

### Production deployment

For a Windows lab PC that runs this service 24/7 (and possibly other device services on the same PC), follow the canonical install recipe in the monorepo:

**[`docs/DEVICE_PC_SETUP.md`](https://github.com/AccelerationConsortium/ac-organic-lab/blob/main/docs/DEVICE_PC_SETUP.md)**

That document covers:

- Installing uv to a system-wide path (`C:\SDL_Tools\uv.exe`) so Windows Services can find it.
- Wrapping `agilent-plateloc-serve` in **NSSM** so it auto-starts on boot, restarts on crash, and writes rotated log files (the systemd-equivalent for Windows).
- Running the service as a real lab user account (not `LocalSystem` — required for the PlateLoc ActiveX profile lookup in `HKCU` to succeed).
- The `update_all.ps1` workflow for keeping multiple device services in sync after a `git push`.
- Troubleshooting the common service-startup failures.

Install uv into `C:\SDL_Tools\uv.exe`:

```powershell
# Run from an elevated PowerShell.
New-Item -ItemType Directory -Force C:\SDL_Tools | Out-Null

# Official uv installer. It installs into the current user's local bin first.
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

# Copy the user-local uv.exe to the stable service path.
# The fallback handles shells where PATH has already been refreshed and uv is
# discoverable via Get-Command.
$uvUser = Join-Path $env:USERPROFILE ".local\bin\uv.exe"
if (!(Test-Path $uvUser)) {
    $uvUser = (Get-Command uv -ErrorAction Stop).Source
}
Copy-Item $uvUser C:\SDL_Tools\uv.exe -Force

# Verify the exact binary NSSM will call later.
C:\SDL_Tools\uv.exe --version
```

Install NSSM:

```powershell
# Preferred on modern Windows.
winget install -e --id NSSM.NSSM

# If winget is unavailable, use Chocolatey or download nssm.exe manually:
# choco install nssm -y
# https://nssm.cc/download
```

Quick version, for the impatient:

```powershell
# As Administrator, after installing uv to C:\SDL_Tools\uv.exe and NSSM:
New-Item -ItemType Directory -Force C:\Users\sdl2\Projects | Out-Null
New-Item -ItemType Directory -Force C:\SDL_Logs            | Out-Null

cd C:\Users\sdl2\Projects
git clone https://github.com/cyrilcaoyang/agilent-plateloc-server.git
cd C:\Users\sdl2\Projects\agilent-plateloc-server
copy config.example.toml config.toml ; notepad config.toml
C:\SDL_Tools\uv.exe sync --extra api

nssm install plateloc C:\SDL_Tools\uv.exe `
    run --project C:\Users\sdl2\Projects\agilent-plateloc-server --extra api agilent-plateloc-serve
nssm set plateloc AppDirectory  C:\Users\sdl2\Projects\agilent-plateloc-server
nssm set plateloc AppStdout     C:\SDL_Logs\plateloc.out.log
nssm set plateloc AppStderr     C:\SDL_Logs\plateloc.err.log
nssm set plateloc AppExit Default Restart
nssm set plateloc ObjectName    ".\labuser" "<password>"
nssm start plateloc
```

After the service is up, register the device in the dashboard's `equipment.yaml` with `adapter: http` and `protocol: "1.2"` (see `docs/STATUS_SPEC.md` §9).

## 32-bit Python Requirement

The Agilent `AgilentPlateLoc.dll` ActiveX control is a **32-bit** COM component.  
If your main Python is 64-bit (which is typical), this driver automatically launches a **32-bit Python subprocess** to host the COM object and communicates with it over JSON pipes — you don't need to change your main Python.

Install a 32-bit Python alongside your main one, then install `pywin32` into that 32-bit runtime:

```powershell
# Check which Python runtimes the launcher can see.
py -0

# This project is currently set up with Python 3.13 (32-bit):
py -3.13-32 -m pip install pywin32
py -3.13-32 -c "import win32com.client, pythoncom; print('pywin32 ok')"

# Option A — Python.org installer
# Download the 32-bit (x86) installer from https://www.python.org/downloads/
# During install, check "Add to PATH" is OFF (to avoid conflicts)
# Then install pywin32 into it, adjusting the selector to match `py -0`.
# For example, if `py -0` shows "-V:3.10-32":
py -3.10-32 -m pip install pywin32

# Option B — winget
winget install Python.Python.3.10 --architecture x86
```

The exact Python version is less important than the architecture: the PlateLoc ActiveX control requires **32-bit Python with `pywin32` installed**. The driver auto-detects 32-bit Python via the `py` launcher (`py -3-32`). You can also pass the path explicitly:

```python
sealer = PlateLoc(python32_path=r"C:\Python310-32\python.exe")
```

## First-Time Profile Setup (Administrator Required)

The PlateLoc ActiveX control stores profiles in a protected registry location.
You **must run as Administrator** the first time to create / edit a profile.

```powershell
# Open an elevated PowerShell:
#   • Press Win+X → select "Windows Terminal (Admin)" or "PowerShell (Admin)"
#   • Or: press Win, type "powershell", right-click → "Run as administrator"

cd path\to\agilent-plateloc-server
.venv\Scripts\python.exe -c "
from agilent_plateloc_server import PlateLoc
s = PlateLoc()          # uses com_port from config.toml
s._create_com_object()
s.show_diags_dialog(modal=True, security_level=0)
s.close()
"
```

In the Diagnostics dialog:

1. Go to the **Profiles** tab
2. Click **Create a new profile** and give it a name (e.g. `MyPlateLoc`)
3. Set **Serial port** to **COM14**
4. Configure startup values (temperature, seal time, etc.)
5. Click **Update this profile** to save
6. Click **OK** to close

> **Note:** You only need Administrator privileges to create or modify profiles.
> Normal operation (`connect`, `start_cycle`, etc.) works without elevation.

## Quick Start

After a profile exists:

```python
from agilent_plateloc_server import PlateLoc

with PlateLoc() as sealer:           # reads com_port from config.toml
    sealer.connect()                  # reads profile from config.toml

    # Configure
    sealer.set_sealing_temperature(170)   # 20–235 °C
    sealer.set_sealing_time(3.0)          # 0.5–12.0 s

    # Read
    print("Hot plate temp:", sealer.get_actual_temperature(), "°C")
    print("Firmware:", sealer.get_firmware_version())
    print("Cycle count:", sealer.get_cycle_count())

    # Seal
    sealer.start_cycle()
```

## API Reference

### Connection

| Method | Description |
|---|---|
| `PlateLoc(com_port, ...)` | Create a driver instance |
| `connect(profile)` | Initialize and connect using a named profile |
| `close()` | Disconnect and release resources |
| `enumerate_profiles()` | List available profile names |
| `show_diags_dialog(modal, security_level)` | Open the Diagnostics / profile editor dialog |

### Sealing

| Method | Description |
|---|---|
| `set_sealing_temperature(°C)` | Set temperature (20–235 °C) |
| `set_sealing_time(seconds)` | Set cycle duration (0.5–12.0 s) |
| `start_cycle()` | Start a seal cycle |
| `stop_cycle()` | Stop a running cycle |
| `apply_seal()` | Apply seal and keep door closed |

### Stage Control

| Method | Description |
|---|---|
| `move_stage_in()` | Move plate stage into the sealing chamber |
| `move_stage_out()` | Move plate stage out of the sealing chamber |

### Readings

| Method | Returns |
|---|---|
| `get_actual_temperature()` | Current hot plate temperature (°C) |
| `get_sealing_temperature()` | Configured sealing set-point (°C) |
| `get_sealing_time()` | Configured seal duration (s) |
| `get_cycle_count()` | Total seal cycles performed (odometer) |
| `get_firmware_version()` | Firmware version string |
| `get_version()` | ActiveX control version string |
| `get_last_error()` | Last error description |

### Error Handling

| Method | Description |
|---|---|
| `abort()` | Abort current task in error state |
| `retry()` | Retry last action after error |
| `ignore_error()` | Ignore last error and proceed |

## REST API

The repo ships a FastAPI service that exposes the driver over HTTP using
the unified lab equipment status spec (v1.1). The dashboard polls this
service every 2-3 seconds; orchestrators acquire a short-lived claim
before issuing writes to `/control/*`.

### Run the service

```powershell
# From an environment that already has the driver deps installed:
pip install -e ".[api]"           # adds fastapi + uvicorn + pydantic

# Production - reads [service] from config.toml
agilent-plateloc-serve

# Or as a module (handy when iterating)
python -m agilent_plateloc_server

# Force dry-run (no hardware) for development on macOS/Linux
python -m agilent_plateloc_server --dry-run --port 8000
```

Configure host/port/dry-run in `config.toml`:

```toml
[service]
host = "0.0.0.0"          # Tailscale-only by ACL
port = 8000
dry_run = false           # true = run without ActiveX/COM (CI, dev)
cors_origins = ["*"]      # tighten if device leaves the Tailnet
startup_connect_timeout_s = 15.0
startup_retry_interval_s = 30.0
                          # v1.5.0: retry a failed boot auto-connect every
                          # N s until the first successful connect (USB
                          # serial can enumerate after the service at
                          # boot). 0 disables. The retry stops permanently
                          # at the first success, so a deliberate
                          # /control/shutdown is never fought.
enforce_claims = true     # v1.1: require X-Claim-Token on /control/*
                          # set false for advisory mode (publishes
                          # claimed_by but doesn't block writes)

[dashboard]
equipment_id = "plateloc"          # MUST match equipment.yaml in the dashboard
equipment_name = "Agilent PlateLoc"
```

### Endpoints

Spec-mandated (always available, no claim required):

| Method | Path             | Returns                                         |
|--------|------------------|-------------------------------------------------|
| GET    | `/`              | `{equipment_id, equipment_name, protocol_version}` |
| GET    | `/health`        | `{status: "healthy"}`                           |
| GET    | `/status`        | Full `EquipmentStatus` envelope (always 200)    |
| GET    | `/openapi.json`  | OpenAPI document (FastAPI auto-generates)       |

Claim protocol (v1.1, no token required to *acquire* a claim):

| Method | Path                  | Body / Headers                                 |
|--------|-----------------------|------------------------------------------------|
| POST   | `/control/claim`      | `{owner, session_id, ttl_s}` -> `ClaimResponse` (or 409 `ClaimRejection`) |
| POST   | `/control/heartbeat`  | header `X-Claim-Token` -> `ClaimResponse` (or 401) |
| POST   | `/control/release`    | header `X-Claim-Token` -> 204 (idempotent)     |

Control (require `X-Claim-Token` matching the live claim, or HTTP 423):

| Method | Path                          | Body                                          |
|--------|-------------------------------|-----------------------------------------------|
| POST   | `/control/startup`            | `{profile?: string}`                          |
| POST   | `/control/shutdown`           | `{}`                                          |
| POST   | `/control/seal/temperature`   | `{temperature_c: int}` (20-235)               |
| POST   | `/control/seal/time`          | `{seconds: float}` (0.5-12.0)                 |
| POST   | `/control/seal/start`         | `{temperature_c?, seconds?}`                  |
| POST   | `/control/seal/stop`          | `{}`                                          |
| POST   | `/control/stage/in`           | `{}`                                          |
| POST   | `/control/stage/out`          | `{}`                                          |

Control endpoints return **423 Locked** when no/wrong `X-Claim-Token` is
provided (with `claimed_by` in the body so the caller can see who holds
the device), **409 Conflict** if the driver isn't connected yet
(operator should hit `/control/startup` first), **422** for out-of-range
parameters, and **503** if connect itself fails.

`POST /control/seal/start` is **synchronous**: the ActiveX control runs in
blocking mode, so the request returns when the physical cycle has finished
(0.5–12 s, per `[film]`/`seconds`). While a cycle is in flight the device
returns **409 Conflict** to anything that would start a second run or move
the carriage — `seal/start`, `stage/in`, `stage/out`, `seal/temperature`,
`seal/time` — exactly matching what `allowed_actions` advertises. `seal/stop`
and `shutdown` stay available.

The `EquipmentStatus` envelope additionally includes:

* **`allowed_actions`** — a flat list of skill names the device will
  currently honour on `/control/*`. Authoritative; the SDK prefers this
  over its own catalog `requires_states` whenever non-empty.
* **`details.claimed_by`** — `{session_id, owner, expires_at}` while a
  claim is held; absent when unclaimed.
* **`activity` / `activity_since`** (v1.2) — `"running"` while a seal cycle
  executes, and the instant that span began.
* **`metrics.cycles_total`** (v1.2) — the instrument's lifetime seal-cycle
  odometer under the spec's reserved key.
* **`details.readings_as_of`** — present only while a cycle owns the COM
  channel, marking the instrument values as the last observation rather than
  a fresh read (the metric timestamps carry the same instant).

### Quick check

```bash
# Probe + health (no claim required)
curl http://plateloc-pc:8000/
curl http://plateloc-pc:8000/health

# Full status snapshot
curl http://plateloc-pc:8000/status | jq

# Acquire a claim, then issue control writes
TOKEN=$(curl -sX POST http://plateloc-pc:8000/control/claim \
  -H 'Content-Type: application/json' \
  -d '{"owner": "alice@cli", "session_id": "demo-1", "ttl_s": 60}' \
  | jq -r .claim_token)

curl -X POST http://plateloc-pc:8000/control/startup \
  -H "X-Claim-Token: $TOKEN" \
  -H 'Content-Type: application/json' -d '{"profile": "default"}'

curl -X POST http://plateloc-pc:8000/control/seal/start \
  -H "X-Claim-Token: $TOKEN" \
  -H 'Content-Type: application/json' \
  -d '{"temperature_c": 170, "seconds": 3.0}'

# Heartbeat every ~heartbeat_interval_s while you still need the device,
# then release on exit.
curl -X POST http://plateloc-pc:8000/control/heartbeat \
  -H "X-Claim-Token: $TOKEN"
curl -X POST http://plateloc-pc:8000/control/release \
  -H "X-Claim-Token: $TOKEN"
```

The Python SDK (`lab_skills.ClaimManager`) handles the
acquire/heartbeat/release loop automatically; raw `curl` is only useful
for one-off operator probing.

### Spec conformance notes

* `GET /status` is **side-effect-free** — polling it never moves the
  stage, never fires a cycle, and never re-initialises the driver.
* `GET /status` always returns **HTTP 200** when the process is alive.
  Hardware-not-yet-initialised is reported as
  `equipment_status: requires_init` with `required_actions: ["startup"]`.
* The `equipment_id` in `/status` matches the `id` in the dashboard's
  `equipment.yaml`. Do not change it without coordinating with the
  dashboard repo.
* No `equipment_ip` / `equipment_tailscale` self-discovery — the
  dashboard registry is the single source of truth for "where to reach
  this device".
* `models.py` no longer vendors the contract: it re-exports the shared
  `sdl-lab-contract` types (pinned to tag `v1.2.0`, the package version
  tracks the spec revision) and keeps only `ClaimRequest`, which this device
  validates more strictly than the wire contract requires.
* `protocol_version` is `"1.2"` on both `/` and `/status`.

Reference snapshots live in `tests/fixtures/status_*.json` covering
`requires_init`, `ready` (+ `ready_claimed`, and the three
`seal.start`-blocked shapes), `busy`, `degraded_running`, `dry_run`, and the
`last_error` taxonomy example. They are regenerated by `pytest` and committed
so reviewers can eyeball schema changes; `test_status_v12.py` also validates
every one of them against `sdl_lab_contract.EquipmentStatus` and the §2.3
consistency invariants.

### Activity and cycle accounting (v1.2)

`equipment_status` answers "is the sealer healthy?"; `activity` answers "is it
sealing right now?" (STATUS_SPEC §2.3). They are independent: a heater-readback
fault mid-cycle reports `degraded` **and** `activity: "running"` — see
`tests/fixtures/status_degraded_running.json`.

| | |
|---|---|
| **Primary operation** | one seal cycle: the span of the blocking `StartCycle` COM call, which is also the span of the `POST /control/seal/start` request |
| **Observed from** | the seal-cycle state machine (`_busy_state`), never derived from `equipment_status`. The Agilent control exposes no "cycle in progress" query, so this is command-tracked — the same mechanism the stage position uses |
| **`activity_since`** | the instant `activity` last *changed*; repeated polls of an unchanged activity do not move it |
| **Invariants** | `busy` ⇒ `running`; `ready` / `requires_init` / `e_stop` ⇒ `idle`; `degraded` ⇒ whichever is true |
| **`allowed_actions`** | while `running`, only `shutdown` + `seal.stop` (no second cycle, no carriage move under a lowered press); the matching `/control/*` refusal is a 409 |

> ⚠️ Like the stage position, this state is **in-memory**. A process restart
> during a cycle reports `idle`; the window is one cycle (≤ 12 s).

**Why `cycles_total` matters here.** A seal cycle is seconds long and the
dashboard aggregator polls at 60 s, so most cycles begin and end between two
polls — they are not undercounted, they are missed outright (§2.3.1).
`metrics["cycles_total"]` mirrors the instrument's **lifetime odometer**, so a
reader recovers what it slept through from the poll-to-poll delta, and the
monotonic semantics hold by construction (the hardware counter never resets;
`cycle_count` is retained unchanged for existing readers and reports the same
number).

**A poll during a cycle answers immediately.** `/status` holds the state lock
only long enough to snapshot in-memory state, then reads the instrument
outside it under a separate COM-channel lock. While a cycle owns that channel
the snapshot serves the last observation (flagged with
`details.readings_as_of`) instead of queueing. Before v1.2 the state lock was
held across the whole blocking call, so a poll issued mid-cycle returned only
after the cycle ended — meaning no reader could ever observe the device
sealing.

### Safety interlocks

This service participates in the **four-layer interlock model** described
in [`docs/INTERLOCKS.md`](https://github.com/AccelerationConsortium/ac-organic-lab/blob/main/docs/INTERLOCKS.md)
in the `ac-organic-lab` monorepo (hardware limits, device state machine,
skill preconditions, and project plan interlocks). The PlateLoc owns
**layer 1** — the device itself refuses unsafe commands.

#### Temperature band — refuses `/control/seal/start` when not at setpoint

| | |
|---|---|
| **Rule** | `abs(actual_temperature − setpoint_temperature) ≤ temperature_tolerance_c` |
| **Tolerance source** | `[film].temperature_tolerance_c` in `config.toml` (default `2` °C). The same value is published on `/status` as `details.temperature_tolerance_c` — there is one source of truth. |
| **HTTP response when violated** | `412 Precondition Failed` with a top-level JSON body and (when known) a `Retry-After` header in seconds. |
| **Config flag** | `[service].enforce_temp_interlock` (default `true`). |
| **Override** | Set `enforce_temp_interlock = false` only for emergency cases like cold calibration. Disabling it restores the pre-interlock failure mode where a seal cycle started below setpoint produces an underspec'd seal and a downstream pneumatic fault on the press. |

Response body on a 412:

```json
{
  "detail": "Temperature outside seal band",
  "actual_c": 150.0,
  "setpoint_c": 170.0,
  "tolerance_c": 2.0,
  "retry_after_s": 21
}
```

`retry_after_s` is a conservative best-effort estimate (heating
~1 °C/s, cooling ~0.3 °C/s); poll `/status` and inspect
`components.heater.state` for the authoritative "ready to seal" signal
rather than burn-polling on the retry hint. When the device cannot
read `actual_temperature` or `setpoint_temperature` at all the
interlock still refuses (with `actual_c` / `setpoint_c` set to `null`
and `retry_after_s` omitted) — refusing safely beats guessing.

The check is enforced in `PlateLocService.start_cycle`, *before* the
`StartCycle` COM call, under the same lock that owns the
actual/setpoint reads — so a concurrent `set_sealing_temperature`
cannot race the precondition.

`v1.2.1` extends the interlock to `/status` as well: the
`allowed_actions` list drops `seal.start` whenever the band check
would refuse it (heater heating, heater cooling, or temperatures
unreadable). Both surfaces consult the same
`PlateLocService.evaluate_temperature_interlock` helper, so a
workflow client that trusts `allowed_actions` verbatim no longer
races into a `412`. The 412 path remains authoritative — `allowed_actions`
is advisory.

> The dashboard tile (`PlateSealerTile` in `ac-organic-lab`) also gates
> "Seal start" on the same band. That is a UX safety net; this
> server-side interlock is the authoritative one — workflows calling
> through the SDK or `curl` on the Tailnet would otherwise bypass the
> tile check.

#### Stage interlock — refuses `/control/seal/start` when the carriage isn't loaded (v1.3.0)

| | |
|---|---|
| **Rule** | `components.stage.state == "in"` |
| **HTTP response when violated** | `412 Precondition Failed` with a top-level JSON body. **No `Retry-After`** — recovery is operator-driven, not time-based. |
| **Config flag** | `[service].enforce_stage_interlock` (default `true`). Independent of `enforce_temp_interlock`. |
| **Override** | Set `enforce_stage_interlock = false` only for emergency overrides. Running with it off restores the failure mode where a seal cycle starts with the carriage extended — wasted hot air, risk of film damage, no actual seal. |

Response body on a 412:

```json
{
  "detail": "Stage not loaded",
  "stage_state": "out",
  "required": "in"
}
```

`stage_state` is `"out"` or `"unknown"`. When BOTH the stage and temperature interlocks would refuse the call, the stage refusal lands first (faster for the operator to fix — one click vs. a temperature ramp). The temperature body is not surfaced in that case; resolve the stage first, then re-POST.

**In-memory, command-tracked state.** The Agilent COM API exposes no stage-position query, so v1.3.0 tracks position by remembering the last commanded direction. The transitions are:

| Trigger                                           | New state |
|---------------------------------------------------|-----------|
| Process startup (in-memory state initialised)     | `unknown` |
| `POST /control/stage/in` returns 200              | `in`      |
| `POST /control/stage/out` returns 200             | `out`     |
| `POST /control/stage/{in,out}` returns 4xx / 5xx  | `unknown` |
| `POST /control/seal/start` returns 200            | `in`      |
| `POST /control/seal/start` returns 412 pre-flight | unchanged |
| `POST /control/seal/start` fails mid-cycle (5xx)  | `unknown` |
| `POST /control/shutdown` returns 200              | `unknown` |
| Driver disconnect / error from COM                | `unknown` |

> ⚠️ **State is in-memory only.** After an NSSM restart (or any process restart) the carriage position is `unknown` and the operator must explicitly `POST /control/stage/in` (or `out`) before the device will accept `/control/seal/start`. The contract is deliberate: the plate may have been moved manually while the service was down, so any persisted state would be a lie. The dashboard tile renders `seal.start` as disabled until the carriage is homed.

**Stage move dedup.** When `stage.state == "in"`, `/status.allowed_actions` omits `stage.in` (no-op direction); same for `stage.out`. The `POST` itself is still accepted as a 200 no-op — the asymmetry vs. `seal.start` is intentional: a redundant stage move is harmless; sealing without a plate is wasted hot air.

`v1.3.0` also extends the existing v1.2.1 pattern: both surfaces (`/status.allowed_actions` and the 412 path) consult the same `PlateLocService.evaluate_stage_interlock` helper, so a workflow client trusting `allowed_actions` verbatim cannot race into a 412 the device would have refused.

#### Health interlock — refuses `/control/seal/start` while a failure is uncleared (v1.4.0)

| | |
|---|---|
| **Rule** | no `last_error` inside the 60 s recent-failure window (the same window that puts the device in `equipment_status: "error"`) |
| **HTTP response when violated** | `412 Precondition Failed` with a top-level JSON body plus `Retry-After` — recovery is time-bounded, and §6.4's auto-clear makes it immediate |
| **Config flag** | none. §2.2 requires it: a device must not start a normal run while it knows of an active fault |

Response body on a 412:

```json
{
  "detail": "Recent operational failure not cleared",
  "last_error_code": "low_air_pressure",
  "last_error_message": "StartCycle returned error code -2147221503 (driver: Low Air Pressure Error)",
  "retry_after_s": 47
}
```

Clearing it is the ordinary §6.4 path: the first 2xx from any operational
endpoint drops `last_error`, and the run reappears on both surfaces at once.
In practice that is the `POST /control/stage/in` an operator would issue
anyway — a mid-cycle failure pessimizes the carriage to `unknown`, so the
stage interlock demands a re-home regardless.

**What this replaced, and why.** Through v1.3.2 the `error` and `degraded`
states collapsed `allowed_actions` to `["shutdown"]` while `/control/*` still
honoured everything the interlocks allowed. That was wrong twice over. It was
a §6.2 violation in the withholding direction — `/status` omitted actions the
device would perform — and, worse, it withheld exactly the actions an operator
needs after a fault: when the 2026-07-15 bench run failed on low air pressure
with a plate in a hot chamber, the device advertised no way to retract the
carriage. v1.4.0 keeps the recovery and diagnostic actions listed in `error` /
`degraded` (§2.2 explicitly permits this) and moves the run's gate into this
interlock, where both surfaces share one helper.

Note the division of labour among the three: `degraded` from a readback that
sealing does not depend on (e.g. the seal-time query) leaves `seal.start`
available, because §2.2's "safe, useful subset" still holds; a heater readback
that fails takes it away through the **temperature** interlock's fail-closed
path, not through the state.

#### `last_error` clears on the next successful operational action

The structured `last_error` block on `/status` reports the most recent
operational failure (driver fault, missing profile, Low Air Pressure
exception, etc.). Prior to `v1.2.1` it stuck around until the process
restarted — operators reading the dashboard would still see a stale
Low Air Pressure error hours after the device had recovered.

`v1.2.1` clears `last_error` on the **first 2xx response** from an
operational endpoint that follows a failure. The policy is enforced at
the API layer (`api.py`) via `service.clear_last_error_on_success()`:

| Endpoint                       | Clears on 2xx? | Notes |
|--------------------------------|----------------|-------|
| `/control/startup`             | yes            | including the already-connected fast path |
| `/control/shutdown`            | yes            |       |
| `/control/seal/start`          | yes            | only on 2xx; a 412 refusal keeps `last_error` |
| `/control/seal/stop`           | yes            |       |
| `/control/seal/temperature`    | yes            |       |
| `/control/seal/time`           | yes            |       |
| `/control/stage/in`            | yes            |       |
| `/control/stage/out`           | yes            |       |
| `/control/claim`               | no             | claim infrastructure, not operational progress |
| `/control/heartbeat`           | no             | same |
| `/control/release`             | no             | same |
| `/`, `/health`, `/status`      | no             | read-only — must not mutate state |

Implementation note: the clear is invoked *after* every service call
in the endpoint has succeeded but *before* the response body is built,
so a multi-step endpoint (e.g. `seal/start` setting time then refusing
on the band check) does not partially clear.

#### `last_error.code` taxonomy (v1.3.1)

`last_error.code` was a free-form string through v1.3.0 (typically the
failing method name). `v1.3.1` promotes it to a **closed enum** so
dashboards branch on `code` and never have to regex-match on
`message`. The wire shape is unchanged (`ErrorInfo.code` is still
`str | None`); the contract is internal validation — `set_last_error`
rejects any value outside the table.

| `code`              | When it fires                                                                                                 | Dashboard recovery hint                          |
|---------------------|---------------------------------------------------------------------------------------------------------------|--------------------------------------------------|
| `low_air_pressure`  | Driver returns "Low Air Pressure" — the lab air supply dropped below the press requirement.                  | "Check air supply / regulator pressure."         |
| `no_plate`          | Driver returns "No Plate In Holder" — a seal cycle started with an empty stage.                               | "Load a plate before sealing."                      |
| `vacuum_error`      | Driver returns "Hot Plate Vacuum Error" — the cycle couldn't draw/hold vacuum (missing/failed seal film, plate mis-seated, or vacuum fault). | "Check seal film and plate seating; verify vacuum." |
| `com_init_failed`   | Startup couldn't reach the physical sealer (e.g. powered off, serial cable disconnected, COM port busy).      | "Check power and serial cable; restart device."  |
| `com_timeout`       | A COM call timed out without a specific driver error code.                                                    | "Driver unresponsive — restart device service."  |
| `com_other`         | Catch-all for driver errors we don't yet classify. A repeated failure landing here is the cue to add a code. | "Unhandled driver fault — file a bug."           |
| `heater_overtemp`   | Driver reports the heater exceeded its safety limit.                                                          | "Heater overtemperature — service required."     |
| `heater_undertemp`  | Heater failed to reach setpoint after the expected ramp window.                                              | "Heater not reaching setpoint — service required."|
| `profile_not_found` | `Initialize()` was called with a profile name not configured in the Diagnostics dialog.                       | "Open Diagnostics dialog and create profile."    |
| `stage_jam`         | A stage move command failed in a way that wasn't simply "stage didn't move" (e.g. driver reports the press is down). | "Check carriage path; recover via Diagnostics."  |
| `process_internal`  | A Python type error (KeyError, AttributeError, etc.) bubbled up — software bug, not a driver fault.          | "Service bug — file an issue."                   |

`v1.4.0` note: a failing instrument *readback* (which drives
`equipment_status: "degraded"`) is now also surfaced as a `last_error` with
`severity: "warning"` and a classified `code`, instead of reaching clients
only as free text in `message`. It is a warning, not an error: §2.2 already
carries the safety consequence in the top-level state, and an operational
failure always takes precedence over a synthesized one.

The classifier (`PlateLocService._classify_error`, text rules shared with
`_classify_error_text`) inspects the
failing method name, the exception, and the driver's
`get_last_error()` detail string, in that order of specificity:
Python type errors first (`process_internal`), then text-based driver
matches (`low_air_pressure`, `no_plate`, `vacuum_error`, `heater_*`),
then `com_timeout`, then
context fallbacks (`stage_jam` for stage moves, `com_init_failed` for
startup), then `com_other` as the default.

The `message` field still carries the driver's free-form text
verbatim — codes classify, messages preserve fidelity. Auto-clear
(see above) drops the entire `last_error` block together, so a
dashboard never sees a partial state where `code` and `message`
disagree.

### Running on the device PC

The PlateLoc PC is a Windows machine on the lab Tailnet. Recommended
process supervisor: NSSM (or the Windows Task Scheduler with
`agilent-plateloc-serve` set to "run whether user is logged on or not").
On Linux for CI/dev, a `systemd` unit pointing at
`agilent-plateloc-serve --dry-run` is enough.

## Project Structure

```
agilent-plateloc-server/
├── README.md
├── pyproject.toml
├── config.example.toml          # Template — copy to config.toml
├── config.toml                  # Your local settings (gitignored)
├── demo.py                      # Demonstration script
├── src/
│   └── agilent_plateloc_server/
│       ├── __init__.py          # Package entry point
│       ├── __main__.py          # CLI: `python -m agilent_plateloc_server`
│       ├── plateloc.py          # Main driver class (ActiveX/COM)
│       ├── _com_server.py       # 32-bit COM surrogate (internal)
│       ├── config.py            # Config loader (reads config.toml)
│       ├── models.py            # Re-exports sdl-lab-contract (spec v1.2)
│       ├── claims.py            # v1.1 ClaimStore (acquire/heartbeat/release)
│       ├── service.py           # PlateLocService - state + locking + dry-run
│       └── api.py               # FastAPI app (spec + claim + control endpoints)
└── tests/
    ├── conftest.py              # TestClient fixtures (dry-run, claimed) + scrubber
    ├── test_api.py              # Spec conformance tests + fixture writer
    ├── test_claims.py           # v1.1 claim protocol conformance tests
    ├── test_status_v12.py       # v1.2 activity / cycles_total / §6.2 agreement
    └── fixtures/
        ├── status_dry_run.json
        ├── status_ready.json                    # stage homed in, seal.start present
        ├── status_ready_claimed.json
        ├── status_ready_stage_unknown.json      # v1.3.0: after startup, before homing
        ├── status_ready_stage_out.json          # v1.3.0: carriage extended
        ├── status_ready_heating.json            # v1.2.1: temp gate blocks seal.start
        ├── status_ready_mid_cycle_failure.json  # v1.3.0: last_error set, stage unknown
        ├── status_last_error_low_air_pressure.json  # v1.3.1: taxonomy wire example
        ├── status_busy.json                     # v1.4.0: busy + activity running
        ├── status_degraded_running.json         # v1.4.0: fault does not hide the cycle
        └── status_requires_init.json
```

## Troubleshooting

### "Class not registered" error

The ActiveX DLL is 32-bit. Make sure you have 32-bit Python with `pywin32` installed:

```powershell
py -3-32 -c "import win32com.client; print('OK')"
```

### "Failed to create COM object" error

Make sure the VWorks ActiveX Controls are installed and registered:

```powershell
# Re-register (run as Administrator)
cd "C:\Program Files (x86)\Agilent Technologies\VWorks ActiveX Controls"
.\registerALL.bat
```

### No profiles available / "Unable to save the profile settings"

Profile management requires **Administrator** privileges. See [First-Time Profile Setup](#first-time-profile-setup-administrator-required) above.

### "Communication failed - Could not open"

The ActiveX control cannot open the serial port. Check:

1. **PlateLoc is powered on** and the serial cable is connected
2. **No other process** has the port open — kill stale Python processes:
   ```powershell
   Get-Process python* | Stop-Process -Force
   ```
3. **COM port is correct** in the profile — verify in Device Manager (Ports → COM & LPT) and re-open the Diagnostics dialog as Administrator to fix if needed

## Configuration

All instrument-specific settings live in `config.toml` (gitignored).  
Copy the template and edit:

```powershell
copy config.example.toml config.toml
```

See `config.example.toml` for all available keys and their defaults.

### Seal parameter settings

Runnable seal defaults are stored in `parameters.json`. The demo uses this file to let the operator select:

1. Seal type
2. Exact plate type
3. Default temperature / time, with a required confirm-or-override prompt

The structure looks like:

```json
{
  "seal_types": [
    {
      "name": "Agilent Thin Clear Pierceable Film",
      "plates": [
        {
          "name": "8R/12C PP Round Well Spherical Bottom (14mm)",
          "temperature_c": 130,
          "time_s": 3.0
        },
        {
          "name": "8R/12C PP Square Well Flat Bottom (19mm)",
          "temperature_c": 140,
          "time_s": 6.0
        }
      ]
    }
  ]
}
```

In `config.toml`, keep instrument settings and temperature wait behavior:

```toml
[film]
temperature_tolerance_c = 2
heat_timeout_s = 120
```

The demo configures the PlateLoc with the confirmed temperature/time, waits until the actual plate temperature is within tolerance of the requested temperature, and only then prompts the operator to press ENTER to start the seal cycle.

`film_settings.json` is still kept as catalog/reference data derived from Agilent’s film selection guide, but `parameters.json` is the source used by the demo workflow.

## Appendix: Alternative install via conda

The recommended path is uv (see [Installation](#installation) above). If your team is already standardised on Anaconda and you'd rather not introduce a second tool, the following also works:

```powershell
conda create -n plateloc python=3.10 -y
conda activate plateloc
pip install -e ".[api,dev]"          # or just .[api] for runtime-only
agilent-plateloc-serve
```

Caveats:

- No `uv.lock`-equivalent: `pip install` resolves PyPI fresh each time, so two installs on different days may pick different transitive versions.
- The NSSM service wrapper is fiddlier with conda — it has to invoke `cmd.exe /c "conda activate plateloc && agilent-plateloc-serve"`, which has caused service-startup races in the field. With uv, NSSM points at `C:\SDL_Tools\uv.exe run --project ...` and there is no activation step.
- The 32-bit Python sub-process for the ActiveX control is unaffected by either choice — that runtime is installed via `py -3.13-32` and lives outside the Python environment manager.

For a multi-device PC running several services 24/7, the uv path in `docs/DEVICE_PC_SETUP.md` is meaningfully simpler and is what the rest of the lab uses.

## Legal / Licensing

- **Intended use**: This package is provided for **research and internal evaluation only**.  
  For any **commercial** or regulated use, you must contact **Agilent Technologies** to obtain appropriate licenses and approvals.
- **ActiveX software**: The PlateLoc **ActiveX / VWorks controls are proprietary Agilent software** and **must be obtained and licensed from Agilent**.  
  This project does **not** distribute those components and is not a replacement for any Agilent license.
- **No affiliation**: This project is an **independent, unofficial** integration helper and is **not affiliated with, endorsed by, or supported by Agilent Technologies**.
- **No warranty / misuse**: The author provides this software **“as is”, without warranty of any kind** and **waives any responsibility for damage, injury, or misuse** arising from its use.  
  You are solely responsible for ensuring safe operation of equipment and compliance with all applicable laws, regulations, and vendor licenses.

### Open‑source license choice

The project currently uses the **MIT License**, which is a simple, permissive license that:

- Allows others to use, modify, and redistribute the code (including commercially),
- While including a strong **“no warranty / no liability”** clause that matches the disclaimer above.

If you want to stay open‑source, **MIT is a good fit here**.  
If you instead want to **legally forbid commercial use of *this driver itself***, you would need a **custom non‑commercial license**, which would no longer be an OSI‑approved open‑source license.
