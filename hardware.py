"""
Hardware data collection module.
Gathers CPU, GPU, RAM, Disk, and Motherboard information using WMI and psutil.
Temperature / SMART data uses a fallback chain:
  LibreHardwareMonitor -> OpenHardwareMonitor -> MSAcpi (WMI) -> smartctl.
"""

import ctypes
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import winreg
from datetime import datetime, timedelta
from typing import Any

import psutil
import pythoncom

try:
    import wmi as wmi_module
except ImportError:
    wmi_module = None

try:
    import clr as _clr
except ImportError:
    _clr = None


# ---------------------------------------------------------------------------
# CPU usage tracking (psutil.cpu_percent uses thread-local state internally,
# so we track cpu_times deltas ourselves at the module level).
# ---------------------------------------------------------------------------

_cpu_lock = threading.Lock()
_prev_cpu_times = None
_prev_per_cpu_times = None


def _calc_busy_pct(t1, t2) -> float:
    """Compute CPU busy % between two cpu_times snapshots."""
    fields = t1._fields
    d = {f: getattr(t2, f) - getattr(t1, f) for f in fields}
    total = sum(d.values())
    if total <= 0:
        return 0.0
    idle = d.get("idle", 0)
    return round((1.0 - idle / total) * 100, 1)


def _get_cpu_usage() -> tuple[float, list[float]]:
    """Return (total_pct, [per_core_pct, ...]).  Thread-safe."""
    global _prev_cpu_times, _prev_per_cpu_times
    cur = psutil.cpu_times()
    cur_per = psutil.cpu_times(percpu=True)
    with _cpu_lock:
        prev, prev_per = _prev_cpu_times, _prev_per_cpu_times
        _prev_cpu_times, _prev_per_cpu_times = cur, cur_per
    if prev is None:
        return 0.0, [0.0] * len(cur_per)
    total_pct = _calc_busy_pct(prev, cur)
    per_core = [
        _calc_busy_pct(p, c) for p, c in zip(prev_per, cur_per)
    ] if prev_per and len(prev_per) == len(cur_per) else []
    return total_pct, per_core


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_resource_path(relative_path: str) -> str:
    """Resolve path to a bundled resource (PyInstaller _MEIPASS or dev dir)."""
    base = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base, relative_path)


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def _wmi_cimv2():
    return wmi_module.WMI(namespace="root/cimv2")


def _wmi_storage():
    return wmi_module.WMI(namespace="root/Microsoft/Windows/Storage")


# ---------------------------------------------------------------------------
# LibreHardwareMonitor — direct DLL access via pythonnet
# ---------------------------------------------------------------------------

_lhm_computer = None
_lhm_init_done = False
_lhm_lock = threading.Lock()


def _init_lhm_direct() -> bool:
    """Initialize LHM Computer object from bundled DLL (call once)."""
    global _lhm_computer, _lhm_init_done
    if _lhm_init_done:
        return _lhm_computer is not None
    _lhm_init_done = True
    if _clr is None:
        return False
    try:
        bin_dir = get_resource_path("bin")
        dll_path = os.path.join(bin_dir, "LibreHardwareMonitorLib.dll")
        if not os.path.isfile(dll_path):
            return False
        if bin_dir not in sys.path:
            sys.path.append(bin_dir)
        _clr.AddReference(dll_path)
        from LibreHardwareMonitor.Hardware import Computer
        c = Computer()
        c.IsCpuEnabled = True
        c.IsGpuEnabled = True
        c.IsStorageEnabled = True
        c.IsMotherboardEnabled = True
        c.Open()
        _lhm_computer = c
        return True
    except Exception:
        return False


def _collect_hw_sensors(hw, out: list[dict[str, Any]]):
    """Collect sensors from a single IHardware instance into *out*."""
    parent_id = str(hw.Identifier)
    for sensor in hw.Sensors:
        val = sensor.Value
        if val is not None:
            out.append({
                "Name": str(sensor.Name),
                "Value": float(val),
                "SensorType": str(sensor.SensorType),
                "Identifier": str(sensor.Identifier),
                "Parent": parent_id,
            })


def _get_lhm_sensors_direct() -> list[dict[str, Any]]:
    """Read all sensors from the LHM Computer object (thread-safe)."""
    if not _lhm_computer:
        return []
    result: list[dict[str, Any]] = []
    with _lhm_lock:
        try:
            for hw in _lhm_computer.Hardware:
                hw.Update()
                _collect_hw_sensors(hw, result)
                for sub in hw.SubHardware:
                    sub.Update()
                    _collect_hw_sensors(sub, result)
        except Exception:
            pass
    return result


def _get_lhm_sensors() -> list[dict[str, Any]]:
    """Try direct DLL first, then LHM/OHM WMI namespaces as fallback."""
    direct = _get_lhm_sensors_direct()
    if direct:
        return direct
    for ns in ("root/LibreHardwareMonitor", "root/OpenHardwareMonitor"):
        try:
            w = wmi_module.WMI(namespace=ns)
            result = []
            for s in w.Sensor():
                result.append({
                    "Name": s.Name,
                    "Value": s.Value,
                    "SensorType": s.SensorType,
                    "Identifier": s.Identifier,
                    "Parent": s.Parent,
                })
            if result:
                return result
        except Exception:
            continue
    return []


def _find_sensors(sensors: list[dict], *, sensor_type: str | None = None,
                  name_contains: str | None = None,
                  parent_contains: str | None = None) -> list[dict]:
    out = []
    for s in sensors:
        if sensor_type and s["SensorType"] != sensor_type:
            continue
        if name_contains and name_contains.lower() not in s["Name"].lower():
            continue
        if parent_contains and parent_contains.lower() not in s["Parent"].lower():
            continue
        out.append(s)
    return out


def _get_acpi_temperatures() -> list[dict[str, Any]]:
    """Fallback: read thermal zones via MSAcpi_ThermalZoneTemperature (admin)."""
    temps: list[dict[str, Any]] = []
    try:
        w = wmi_module.WMI(namespace="root/WMI")
        for tz in w.MSAcpi_ThermalZoneTemperature():
            kelvin_tenths = tz.CurrentTemperature
            celsius = round((kelvin_tenths / 10.0) - 273.15, 1)
            name = tz.InstanceName or "Thermal Zone"
            name = name.split("\\")[-1] if "\\" in name else name
            temps.append({"name": name, "value": celsius})
    except Exception:
        pass
    return temps


# ---------------------------------------------------------------------------
# smartctl helper
# ---------------------------------------------------------------------------

_SMARTCTL_PATHS = [
    get_resource_path(os.path.join("bin", "smartctl.exe")),
    shutil.which("smartctl"),
    r"C:\Program Files\smartmontools\bin\smartctl.exe",
    r"C:\Program Files (x86)\smartmontools\bin\smartctl.exe",
]


def _find_smartctl() -> str | None:
    for p in _SMARTCTL_PATHS:
        if p and os.path.isfile(p):
            return p
    return None


def _smartctl_json(device: str) -> dict | None:
    exe = _find_smartctl()
    if not exe:
        return None
    try:
        r = subprocess.run(
            [exe, "-a", "-j", device],
            capture_output=True, text=True, timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        return json.loads(r.stdout)
    except Exception:
        return None


def _smartctl_scan() -> list[str]:
    exe = _find_smartctl()
    if not exe:
        return []
    try:
        r = subprocess.run(
            [exe, "--scan", "-j"],
            capture_output=True, text=True, timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        data = json.loads(r.stdout)
        return [d["name"] for d in data.get("devices", [])]
    except Exception:
        return []


# ---------------------------------------------------------------------------
# CPU
# ---------------------------------------------------------------------------

def _get_dynamic_cpu_freq() -> float | None:
    """Get real-time CPU frequency via PercentProcessorPerformance.

    Uses ProcessorFrequency * PercentProcessorPerformance / 100 which
    correctly reflects turbo boost and manual overclocking (unlike
    PercentofMaximumFrequency which caps at 100 on many OC setups).
    """
    try:
        w = _wmi_cimv2()
        perf = w.Win32_PerfFormattedData_Counters_ProcessorInformation(Name="_Total")
        for p in perf:
            base = int(p.ProcessorFrequency or 0)
            ppp = int(p.PercentProcessorPerformance or 0)
            if base and ppp:
                return round(base * ppp / 100.0, 1)
    except Exception:
        pass

    try:
        freq = psutil.cpu_freq()
        if freq and freq.current:
            return round(freq.current, 1)
    except Exception:
        pass
    return None


def get_cpu_info() -> dict[str, Any]:
    info: dict[str, Any] = {
        "name": "N/A",
        "hardware_id": "N/A",
        "cores": "N/A",
        "threads": "N/A",
        "base_clock_mhz": "N/A",
        "current_clock_mhz": "N/A",
        "temperatures": [],
        "temp_source": "",
    }
    try:
        w = _wmi_cimv2()
        for cpu in w.Win32_Processor():
            info["name"] = cpu.Name.strip()
            info["hardware_id"] = cpu.DeviceID
            info["cores"] = cpu.NumberOfCores
            info["threads"] = cpu.NumberOfLogicalProcessors
            info["base_clock_mhz"] = int(cpu.MaxClockSpeed or 0) or "N/A"
            break
    except Exception:
        pass

    try:
        dynamic = _get_dynamic_cpu_freq()
        if dynamic:
            info["current_clock_mhz"] = dynamic
    except Exception:
        pass

    # Temperature: LHM/OHM -> ACPI fallback
    try:
        sensors = _get_lhm_sensors()
        temps = _find_sensors(sensors, sensor_type="Temperature", parent_contains="cpu")
        info["temperatures"] = [
            {"name": t["Name"], "value": round(t["Value"], 1)} for t in temps if t["Value"] is not None
        ]
        if info["temperatures"]:
            info["temp_source"] = "LibreHardwareMonitor"
    except Exception:
        pass

    if not info["temperatures"]:
        acpi = _get_acpi_temperatures()
        if acpi:
            info["temperatures"] = acpi
            info["temp_source"] = "ACPI Thermal Zone"

    return info


# ---------------------------------------------------------------------------
# GPU
# ---------------------------------------------------------------------------

_PCI_ID_RE = re.compile(r"VEN_([0-9A-Fa-f]{4})&DEV_([0-9A-Fa-f]{4})", re.IGNORECASE)


def _parse_pci_id(pnp_device_id: str) -> str:
    m = _PCI_ID_RE.search(pnp_device_id or "")
    if m:
        return f"{m.group(1).upper()}:{m.group(2).upper()}"
    return "N/A"


_SUBSYS_RE = re.compile(r"SUBSYS_([0-9A-Fa-f]{8})", re.IGNORECASE)

_SUBSYS_VENDOR_MAP: dict[str, str] = {
    "1002": "AMD", "1025": "Acer", "1028": "Dell", "103C": "HP",
    "1043": "ASUS", "1048": "Elsa", "105B": "Foxconn",
    "10B0": "Gainward", "10DE": "NVIDIA",
    "1179": "Toshiba", "1458": "Gigabyte", "1462": "MSI",
    "1569": "Palit", "1682": "XFX", "17AA": "Lenovo",
    "196E": "PNY", "19DA": "Zotac", "1DA2": "Sapphire",
    "3408": "Gainward", "3842": "EVGA", "7377": "Colorful",
    "148C": "PowerColor", "1B4C": "KFA2", "1ACC": "POV",
}


_CHIP_VENDORS = {"10DE", "1002"}


_VEN_RE = re.compile(r"VEN_([0-9A-Fa-f]{4})", re.IGNORECASE)


def _parse_subsys_vendor(pnp_device_id: str) -> str:
    m = _SUBSYS_RE.search(pnp_device_id or "")
    if not m:
        return ""
    raw = m.group(1).upper()
    hi = raw[:4]
    lo = raw[4:]
    name_lo = _SUBSYS_VENDOR_MAP.get(lo, "")
    name_hi = _SUBSYS_VENDOR_MAP.get(hi, "")
    if name_lo and lo not in _CHIP_VENDORS:
        return name_lo
    if name_hi and hi not in _CHIP_VENDORS:
        return name_hi
    # SUBSYS vendor == chip vendor → reference design board
    ven = _VEN_RE.search(pnp_device_id or "")
    chip_id = ven.group(1).upper() if ven else ""
    subsys_vendor_id = lo if name_lo else hi
    vendor_name = name_lo or name_hi
    if vendor_name and subsys_vendor_id == chip_id:
        return f"{vendor_name} (Reference)"
    return vendor_name


def get_gpu_info() -> list[dict[str, Any]]:
    gpus: list[dict[str, Any]] = []
    try:
        w = _wmi_cimv2()
        for vc in w.Win32_VideoController():
            vram_bytes = int(vc.AdapterRAM or 0)
            if vram_bytes < 0:
                vram_bytes = vram_bytes & 0xFFFFFFFF
            vram_gb = round(vram_bytes / (1024 ** 3), 1) if vram_bytes else "N/A"
            raw_name = (vc.Name or "N/A").strip()
            vendor = _parse_subsys_vendor(vc.PNPDeviceID)
            driver_ver = (vc.DriverVersion or "N/A").strip()
            driver_date_raw = vc.DriverDate or ""
            driver_date = "N/A"
            if driver_date_raw and len(driver_date_raw) >= 8:
                try:
                    driver_date = f"{driver_date_raw[:4]}-{driver_date_raw[4:6]}-{driver_date_raw[6:8]}"
                except Exception:
                    pass
            gpus.append({
                "name": raw_name,
                "vendor": vendor or "N/A",
                "pci_device_id": _parse_pci_id(vc.PNPDeviceID),
                "pnp_device_id": vc.PNPDeviceID or "N/A",
                "vram_gb": vram_gb,
                "temperature": "N/A",
                "driver_version": driver_ver,
                "driver_date": driver_date,
            })
    except Exception:
        gpus.append({
            "name": "Access Denied",
            "vendor": "N/A",
            "pci_device_id": "N/A",
            "pnp_device_id": "N/A",
            "vram_gb": "N/A",
            "temperature": "N/A",
            "driver_version": "N/A",
            "driver_date": "N/A",
        })

    try:
        sensors = _get_lhm_sensors()
        gpu_temps = _find_sensors(sensors, sensor_type="Temperature", parent_contains="gpu")
        for i, gpu in enumerate(gpus):
            if i < len(gpu_temps) and gpu_temps[i]["Value"] is not None:
                gpu["temperature"] = round(gpu_temps[i]["Value"], 1)
    except Exception:
        pass

    return gpus


# ---------------------------------------------------------------------------
# RAM
# ---------------------------------------------------------------------------

def _detect_channel_mode(bank_labels: list[str]) -> str:
    """Heuristic: if DIMMs sit in distinct banks/channels -> Dual Channel."""
    if len(bank_labels) < 2:
        return "Single Channel"

    normalized = set()
    for bl in bank_labels:
        bl_lower = bl.lower().strip()
        for token in ("channel a", "channel b", "channel c", "channel d"):
            if token in bl_lower:
                normalized.add(token)
        m = re.search(r"bank\s*(\d+)", bl_lower)
        if m:
            normalized.add(f"bank{m.group(1)}")

    if len(normalized) >= 2:
        return "Dual Channel"
    return "Single Channel"


def get_ram_info() -> dict[str, Any]:
    info: dict[str, Any] = {
        "total_gb": "N/A",
        "available_gb": "N/A",
        "total_slots": "N/A",
        "used_slots": 0,
        "channel_mode": "N/A",
        "sticks": [],
    }

    try:
        vm = psutil.virtual_memory()
        info["total_gb"] = round(vm.total / (1024 ** 3), 1)
        info["available_gb"] = round(vm.available / (1024 ** 3), 1)
    except Exception:
        pass

    bank_labels: list[str] = []
    try:
        w = _wmi_cimv2()
        for mem in w.Win32_PhysicalMemory():
            cap = int(mem.Capacity or 0)
            bank = mem.BankLabel or ""
            bank_labels.append(bank)
            info["sticks"].append({
                "manufacturer": (mem.Manufacturer or "N/A").strip(),
                "serial": (mem.SerialNumber or "N/A").strip(),
                "capacity_gb": round(cap / (1024 ** 3), 1) if cap else "N/A",
                "speed_mhz": mem.Speed or "N/A",
                "bank": bank,
                "locator": mem.DeviceLocator or "N/A",
            })
        info["used_slots"] = len(info["sticks"])
    except Exception:
        pass

    try:
        w = _wmi_cimv2()
        for arr in w.Win32_PhysicalMemoryArray():
            info["total_slots"] = arr.MemoryDevices
            break
    except Exception:
        pass

    info["channel_mode"] = _detect_channel_mode(bank_labels)
    return info


# ---------------------------------------------------------------------------
# Disks
# ---------------------------------------------------------------------------

_MEDIA_TYPE_MAP = {0: "Unspecified", 3: "HDD", 4: "SSD", 5: "SCM"}


def get_disk_info() -> list[dict[str, Any]]:
    disks: list[dict[str, Any]] = []

    wmi_index_to_device: dict[int, str] = {}
    try:
        w = _wmi_cimv2()
        for d in w.Win32_DiskDrive():
            size_bytes = int(d.Size or 0)
            idx = int(d.Index) if d.Index is not None else len(disks)
            wmi_index_to_device[idx] = d.DeviceID or ""
            disks.append({
                "model": (d.Model or "N/A").strip(),
                "serial": (d.SerialNumber or "N/A").strip(),
                "size_gb": round(size_bytes / (1024 ** 3), 1) if size_bytes else "N/A",
                "type": "N/A",
                "temperature": "N/A",
                "power_on_hours": "N/A",
                "health_pct": "N/A",
                "_device_id": d.DeviceID or "",
            })
    except Exception:
        disks.append({
            "model": "Access Denied",
            "serial": "N/A",
            "size_gb": "N/A",
            "type": "N/A",
            "temperature": "N/A",
            "power_on_hours": "N/A",
            "health_pct": "N/A",
            "_device_id": "",
        })

    # Disk type via MSFT_PhysicalDisk
    try:
        ws = _wmi_storage()
        phys = {p.Model.strip(): int(p.MediaType or 0) for p in ws.MSFT_PhysicalDisk()}
        for disk in disks:
            mt = phys.get(disk["model"], 0)
            disk["type"] = _MEDIA_TYPE_MAP.get(mt, "Unknown")
    except Exception:
        pass

    # Temperatures + SMART from LibreHardwareMonitor / OpenHardwareMonitor
    try:
        sensors = _get_lhm_sensors()
        if sensors:
            _enrich_disks_from_lhm(disks, sensors)
    except Exception:
        pass

    # Fallback: smartctl for disks still missing data
    _enrich_disks_from_smartctl(disks)

    # Fallback: WMI FailurePredictStatus for health
    _enrich_disks_failure_predict(disks)

    for disk in disks:
        disk.pop("_device_id", None)

    return disks


def _enrich_disks_from_lhm(disks: list[dict], sensors: list[dict]):
    """Match LHM disk hardware entries to our disk list and pull temps/SMART."""
    disk_keywords = ("hdd", "ssd", "nvme", "disk", "storage")

    disk_hw_ids: list[str] = []
    seen = set()
    for s in sensors:
        parent = s["Parent"].lower()
        if any(kw in parent for kw in disk_keywords) and parent not in seen:
            seen.add(parent)
            disk_hw_ids.append(s["Parent"])

    for i, hw_id in enumerate(disk_hw_ids):
        if i >= len(disks):
            break

        hw_sensors = [s for s in sensors if s["Parent"] == hw_id]

        for s in hw_sensors:
            if s["SensorType"] == "Temperature" and s["Value"] is not None:
                disks[i]["temperature"] = round(s["Value"], 1)
                break

        for s in hw_sensors:
            name_l = s["Name"].lower()
            if "power-on" in name_l or "power on" in name_l:
                if s["Value"] is not None:
                    disks[i]["power_on_hours"] = int(s["Value"])
                break

        for s in hw_sensors:
            name_l = s["Name"].lower()
            if s["SensorType"] == "Level":
                if any(kw in name_l for kw in ("remaining life", "available spare", "life left")):
                    if s["Value"] is not None:
                        disks[i]["health_pct"] = round(s["Value"], 1)
                    break
                if "percentage used" in name_l and s["Value"] is not None:
                    disks[i]["health_pct"] = round(100.0 - s["Value"], 1)
                    break


def _enrich_disks_from_smartctl(disks: list[dict]):
    """Fallback: use smartctl for disks that still have N/A fields."""
    if not _find_smartctl():
        return

    scan_devices = _smartctl_scan()

    for i, disk in enumerate(disks):
        needs_data = (disk["temperature"] == "N/A"
                      or disk["power_on_hours"] == "N/A"
                      or disk["health_pct"] == "N/A")
        if not needs_data:
            continue

        device = scan_devices[i] if i < len(scan_devices) else None
        if not device:
            dev_id = disk.get("_device_id", "")
            if dev_id:
                device = dev_id.replace("\\\\", "\\")
            else:
                continue

        smart = _smartctl_json(device)
        if not smart:
            continue

        if disk["temperature"] == "N/A":
            temp = smart.get("temperature", {}).get("current")
            if temp is not None:
                disk["temperature"] = round(float(temp), 1)

        attrs = {a["name"]: a for a in smart.get("ata_smart_attributes", {}).get("table", [])}

        if disk["power_on_hours"] == "N/A":
            poh_attr = attrs.get("Power_On_Hours")
            if poh_attr:
                disk["power_on_hours"] = int(poh_attr["raw"]["value"])
            else:
                poh_val = smart.get("power_on_time", {}).get("hours")
                if poh_val is not None:
                    disk["power_on_hours"] = int(poh_val)

        if disk["health_pct"] == "N/A":
            if smart.get("smart_status", {}).get("passed") is not None:
                passed = smart["smart_status"]["passed"]
                disk["health_pct"] = "OK" if passed else "FAILING"

            pct_used_attr = attrs.get("Percent_Lifetime_Remain") or attrs.get("Wear_Leveling_Count")
            if pct_used_attr:
                val = int(pct_used_attr.get("value", 0))
                if val:
                    disk["health_pct"] = val


def _enrich_disks_failure_predict(disks: list[dict]):
    """Fallback: WMI MSStorageDriver_FailurePredictStatus for basic health."""
    needs_any = any(d["health_pct"] == "N/A" for d in disks)
    if not needs_any:
        return
    try:
        w = wmi_module.WMI(namespace="root/WMI")
        statuses = list(w.MSStorageDriver_FailurePredictStatus())
        for i, disk in enumerate(disks):
            if disk["health_pct"] != "N/A":
                continue
            if i < len(statuses):
                predict = statuses[i].PredictFailure
                disk["health_pct"] = "FAILING" if predict else "OK"
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Motherboard
# ---------------------------------------------------------------------------

def get_motherboard_info() -> dict[str, Any]:
    info: dict[str, Any] = {
        "manufacturer": "N/A",
        "model": "N/A",
        "bios_version": "N/A",
        "bios_date": "N/A",
    }
    try:
        w = _wmi_cimv2()
        for bb in w.Win32_BaseBoard():
            info["manufacturer"] = (bb.Manufacturer or "N/A").strip()
            info["model"] = (bb.Product or "N/A").strip()
            break
    except Exception:
        pass

    try:
        w = _wmi_cimv2()
        for bios in w.Win32_BIOS():
            info["bios_version"] = bios.SMBIOSBIOSVersion or "N/A"
            raw_date = bios.ReleaseDate or ""
            if raw_date:
                info["bios_date"] = f"{raw_date[:4]}-{raw_date[4:6]}-{raw_date[6:8]}"
            break
    except Exception:
        pass

    return info


# ---------------------------------------------------------------------------
# Aggregate
# ---------------------------------------------------------------------------

def collect_all() -> dict[str, Any]:
    """Collect every hardware category and return a single dict.

    Initializes COM for the current thread so WMI works from
    any thread (background workers, PyInstaller --windowed builds, etc.).
    Also initializes LibreHardwareMonitor DLL on first call.
    """
    pythoncom.CoInitialize()
    try:
        _init_lhm_direct()
        _get_cpu_usage()
        return {
            "cpu": get_cpu_info(),
            "gpu": get_gpu_info(),
            "ram": get_ram_info(),
            "disks": get_disk_info(),
            "motherboard": get_motherboard_info(),
        }
    finally:
        pythoncom.CoUninitialize()


def collect_dynamic() -> dict[str, Any]:
    """Lightweight collection for the 2-second refresh loop.

    Only queries CPU frequency, RAM available, and temperatures.
    """
    pythoncom.CoInitialize()
    try:
        result: dict[str, Any] = {
            "current_clock_mhz": "N/A",
            "cpu_temps": [],
            "gpu_temps": [],
            "ram_available_gb": "N/A",
            "temp_source": "",
            "cpu_usage_total": 0.0,
            "cpu_usage_per_core": [],
            "gpu_load": [],
            "disk_activity": [],
        }

        try:
            w = _wmi_cimv2()
            perf = w.Win32_PerfFormattedData_Counters_ProcessorInformation(Name="_Total")
            for p in perf:
                base = int(p.ProcessorFrequency or 0)
                ppp = int(p.PercentProcessorPerformance or 0)
                if base and ppp:
                    result["current_clock_mhz"] = round(base * ppp / 100.0, 1)
                break
        except Exception:
            pass

        try:
            vm = psutil.virtual_memory()
            result["ram_available_gb"] = round(vm.available / (1024 ** 3), 1)
        except Exception:
            pass

        try:
            total_pct, per_core_pct = _get_cpu_usage()
            result["cpu_usage_total"] = total_pct
            result["cpu_usage_per_core"] = per_core_pct
        except Exception:
            pass

        sensors = _get_lhm_sensors()
        if sensors:
            cpu_t = _find_sensors(sensors, sensor_type="Temperature", parent_contains="cpu")
            result["cpu_temps"] = [
                {"name": t["Name"], "value": round(t["Value"], 1)}
                for t in cpu_t if t["Value"] is not None
            ]
            gpu_t = _find_sensors(sensors, sensor_type="Temperature", parent_contains="gpu")
            result["gpu_temps"] = [
                {"name": t["Name"], "value": round(t["Value"], 1)}
                for t in gpu_t if t["Value"] is not None
            ]
            disk_kw = ("hdd", "ssd", "nvme", "disk", "storage")
            disk_t = [
                s for s in sensors
                if s["SensorType"] == "Temperature"
                and any(kw in s["Parent"].lower() for kw in disk_kw)
                and s["Value"] is not None
            ]
            result["disk_temps"] = [
                {"name": t["Name"], "value": round(t["Value"], 1)}
                for t in disk_t
            ]

            gpu_load = _find_sensors(sensors, sensor_type="Load",
                                     name_contains="GPU Core", parent_contains="gpu")
            result["gpu_load"] = [
                round(s["Value"], 1) for s in gpu_load if s["Value"] is not None
            ]

            disk_kw_load = ("hdd", "ssd", "nvme", "disk", "storage")
            disk_load = [
                s for s in sensors
                if s["SensorType"] == "Load"
                and any(kw in s["Parent"].lower() for kw in disk_kw_load)
                and "total" in s["Name"].lower()
                and s["Value"] is not None
            ]
            result["disk_activity"] = [
                round(s["Value"], 1) for s in disk_load
            ]

            result["temp_source"] = "LHM"

        if not result["cpu_temps"]:
            acpi = _get_acpi_temperatures()
            if acpi:
                result["cpu_temps"] = acpi
                result["temp_source"] = "ACPI"

        result["top_processes"] = get_top_processes()

        return result
    finally:
        pythoncom.CoUninitialize()


def check_lhm_available() -> bool:
    """Check whether LHM is usable — direct DLL or WMI namespace."""
    if _lhm_computer is not None:
        return True
    for ns in ("root/LibreHardwareMonitor", "root/OpenHardwareMonitor"):
        try:
            wmi_module.WMI(namespace=ns)
            return True
        except Exception:
            continue
    return False


def get_top_processes(count: int = 3) -> dict[str, list[dict]]:
    """Return top processes by CPU and RAM usage."""
    try:
        procs = []
        for p in psutil.process_iter(['pid', 'name', 'cpu_percent', 'memory_percent']):
            try:
                info = p.info
                if info['name'] and info['name'] not in ('System Idle Process', 'Idle', ''):
                    procs.append(info)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        top_cpu = sorted(procs, key=lambda x: x.get('cpu_percent', 0) or 0, reverse=True)[:count]
        top_ram = sorted(procs, key=lambda x: x.get('memory_percent', 0) or 0, reverse=True)[:count]
        return {
            "top_cpu": [{"name": p["name"], "pct": round(p.get("cpu_percent", 0) or 0, 1)} for p in top_cpu],
            "top_ram": [{"name": p["name"], "pct": round(p.get("memory_percent", 0) or 0, 1)} for p in top_ram],
        }
    except Exception:
        return {"top_cpu": [], "top_ram": []}


# ---------------------------------------------------------------------------
# Autostart (Windows registry)
# ---------------------------------------------------------------------------

_AUTOSTART_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
_AUTOSTART_NAME = "HardwareDiag"


def get_autostart() -> bool:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _AUTOSTART_KEY, 0, winreg.KEY_READ) as key:
            winreg.QueryValueEx(key, _AUTOSTART_NAME)
            return True
    except (FileNotFoundError, OSError):
        return False


def set_autostart(enable: bool):
    exe = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(sys.argv[0])
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _AUTOSTART_KEY, 0, winreg.KEY_SET_VALUE) as key:
            if enable:
                winreg.SetValueEx(key, _AUTOSTART_NAME, 0, winreg.REG_SZ, f'"{exe}"')
            else:
                try:
                    winreg.DeleteValue(key, _AUTOSTART_NAME)
                except FileNotFoundError:
                    pass
    except OSError:
        pass


def check_smartctl_available() -> bool:
    return _find_smartctl() is not None
