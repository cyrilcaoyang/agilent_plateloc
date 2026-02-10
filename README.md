# Agilent PlateLoc Thermal Microplate Sealer — Python Driver

Python driver for the **Agilent PlateLoc Thermal Microplate Sealer**, communicating through the VWorks ActiveX COM control over a serial (COM) port.

## Prerequisites

- **Windows** (ActiveX is Windows-only)
- **Python 3.10+**
- **Agilent VWorks ActiveX Controls** installed (from the Agilent software CD/UFD)
- **32-bit Python** installed alongside your main Python (the ActiveX DLL is 32-bit — see [32-bit note](#32-bit-python-requirement) below)
- PlateLoc connected via **RS-232 serial** (this project uses **COM14**)

## Installation

### Development — uv environment

[uv](https://docs.astral.sh/uv/) is the recommended tool for fast, reproducible development environments.

```powershell
# Install uv (if you don't have it)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

# Clone / navigate to the project
cd C:\Users\SDL2\Projects\agilent_plateloc

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

Install a 32-bit Python alongside your main one:

```powershell
# Option A — Python launcher (recommended)
# Download the 32-bit (x86) installer from https://www.python.org/downloads/
# During install, check "Add to PATH" is OFF (to avoid conflicts)
# Then install pywin32 into it:
py -3.10-32 -m pip install pywin32

# Option B — winget
winget install Python.Python.3.10 --architecture x86
```

The driver auto-detects 32-bit Python via the `py` launcher (`py -3-32`). You can also pass the path explicitly:

```python
sealer = PlateLoc(com_port="COM14", python32_path=r"C:\Python310-32\python.exe")
```

## First-Time Profile Setup (Administrator Required)

The PlateLoc ActiveX control stores profiles in a protected registry location.
You **must run as Administrator** the first time to create / edit a profile.

```powershell
# Open an elevated PowerShell:
#   • Press Win+X → select "Windows Terminal (Admin)" or "PowerShell (Admin)"
#   • Or: press Win, type "powershell", right-click → "Run as administrator"

cd C:\Users\SDL2\Projects\agilent_plateloc
.venv\Scripts\python.exe -c "
from agilent_plateloc import PlateLoc
s = PlateLoc(com_port='COM14')
s._create_com_object()
s.show_diags_dialog(modal=True, security_level=0)
s.close()
"
```

In the Diagnostics dialog:

1. Go to the **Profiles** tab
2. Click **Create a new profile** and give it a name (e.g. `SDL2_PlateLoc`)
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

with PlateLoc(com_port="COM14") as sealer:
    sealer.connect(profile="my_profile")

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
├── docs/
│   └── VWorks ActiveX Controls 13.1.9 Release Notes.pdf
└── src/
    └── agilent_plateloc/
        ├── __init__.py          # Package entry point, exports PlateLoc
        ├── plateloc.py          # Main driver class
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
