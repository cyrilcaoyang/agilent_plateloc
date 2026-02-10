"""
Agilent PlateLoc Thermal Microplate Sealer - Python Driver
==========================================================

Control the PlateLoc Sealer via its VWorks ActiveX COM interface from Python.

The ActiveX DLL is 32-bit, so this package uses a 32-bit COM surrogate
subprocess to bridge the gap when running under 64-bit Python.

Usage::

    from agilent_plateloc import PlateLoc

    sealer = PlateLoc(com_port="COM14")
    sealer.connect(profile="default")
    sealer.set_sealing_temperature(170)
    sealer.set_sealing_time(3.0)
    sealer.start_cycle()
    temp = sealer.get_actual_temperature()
    sealer.close()
"""

from .plateloc import PlateLoc

__all__ = ["PlateLoc"]
__version__ = "0.1.0"
