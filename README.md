# Agilent PlateLoc Thermal Microplate Sealer — Python Driver

Python driver for the **Agilent PlateLoc Thermal Microplate Sealer**, communicating through the VWorks ActiveX COM control over a serial (COM) port.

## Prerequisites

- **Windows** (ActiveX is Windows-only)
- **Python 3.10+**
- **Agilent VWorks ActiveX Controls** installed (from the Agilent software CD/UFD)
- **32-bit Python** installed alongside your main Python (the ActiveX DLL is 32-bit — see [32-bit note](#32-bit-python-requirement) below)
- PlateLoc connected via **RS-232 serial** (e.g. COM14 — set in `config.toml`)

## Installation

### Development — uv environment

[uv](https://docs.astral.sh/uv/) is the recommended tool for fast, reproducible development environments.

```powershell
# Install uv (if you don't have it)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

# Clone / navigate to the project
cd path\to\agilent_plateloc

# Copy the example config and edit for your setup
copy config.example.toml config.toml
# Edit config.toml — set com_port, profile name, etc.

# Create a virtual environment with Python 3.10 and install with dev deps
uv venv .venv --python 3.10
.venv\Scripts\activate
uv pip install -e ".[dev]"

# Run tests
pytest
```

### Production — conda environment

```powershell
# Create a dedicated conda environment
conda create -n plateloc python=3.10 -y
conda activate plateloc

# Install the package from the local source
pip install .

# Or install in editable mode if you want to iterate
pip install -e .
```

If the package is later published to a private PyPI / Artifactory:

```powershell
conda activate plateloc
pip install agilent-plateloc
```

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

cd path\to\agilent_plateloc
.venv\Scripts\python.exe -c "
from agilent_plateloc import PlateLoc
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
from agilent_plateloc import PlateLoc

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

## Project Structure

```
agilent_plateloc/
├── README.md
├── pyproject.toml
├── config.example.toml          # Template — copy to config.toml
├── config.toml                  # Your local settings (gitignored)
├── demo.py                      # Demonstration script
└── src/
    └── agilent_plateloc/
        ├── __init__.py          # Package entry point
        ├── plateloc.py          # Main driver class
        ├── config.py            # Config loader (reads config.toml)
        └── _com_server.py       # 32-bit COM surrogate (internal)
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
