"""
Diagnostics module for Used PC analysis.
Provides health verdict, red flags detection, deep S.M.A.R.T. analysis,
battery health, stress tests (CPU, Disk, RAM), and PDF report generation.
"""

import ctypes
import ctypes.wintypes
import json
import math
import os
import subprocess
import tempfile
import threading
import time
from typing import Any, Callable

import psutil  # type: ignore
import pythoncom  # type: ignore

from hardware import (  # type: ignore
    _smartctl_json, _smartctl_scan, _find_smartctl,
    _get_lhm_sensors, _find_sensors, _get_cpu_usage,
    get_resource_path,
)


# ---------------------------------------------------------------------------
# S.M.A.R.T. attribute descriptions (for tooltip popups)
# ---------------------------------------------------------------------------

SMART_ATTR_DESC: dict[str, str] = {
    "Raw_Read_Error_Rate":
        "Rate of hardware read errors encountered during disk reads. "
        "On Seagate/Samsung drives, high raw values are NORMAL due to "
        "a different counting method. Only worry if Value < Threshold.",
    "Throughput_Performance":
        "Overall throughput performance of the drive. "
        "Lower values may indicate degradation.",
    "Spin_Up_Time":
        "Average time (in ms) to spin up the platters. "
        "Higher values may indicate bearing wear or power issues.",
    "Start_Stop_Count":
        "Total number of spindle start/stop cycles. "
        "Very high counts can indicate frequent power cycling.",
    "Reallocated_Sector_Ct":
        "CRITICAL: Count of bad sectors that were replaced with spare sectors. "
        "ANY non-zero raw value means the disk surface is physically damaged. "
        "If this number keeps growing, the drive is failing.",
    "Read_Channel_Margin":
        "Margin of a read channel. Vendor-specific diagnostic attribute.",
    "Seek_Error_Rate":
        "Rate of seek errors. On Seagate drives, high raw values are normal "
        "due to vendor-specific counting. Only worry if Value < Threshold.",
    "Seek_Time_Performance":
        "Average efficiency of seek operations. "
        "Declining values can indicate mechanical wear.",
    "Power_On_Hours":
        "Total hours the drive has been powered on. "
        "For used PCs: >20,000h = heavy use, >40,000h = very old drive.",
    "Spin_Retry_Count":
        "Number of failed attempts to spin up the platters. "
        "Non-zero values may indicate motor problems or insufficient power supply.",
    "Power_Cycle_Count":
        "Total power on/off cycles. Shows how many times the drive was restarted.",
    "Reported_Uncorrect":
        "Number of errors that could not be corrected by ECC. "
        "Non-zero values indicate data integrity issues.",
    "High_Fly_Writes":
        "Number of write operations where the head flew too high. "
        "Can cause data integrity issues; indicates vibration or mechanical problems.",
    "Temperature_Celsius":
        "Current drive temperature in Celsius. "
        "Safe range: 25-45°C. Above 55°C accelerates wear.",
    "Hardware_ECC_Recovered":
        "Count of errors recovered by hardware ECC. "
        "High values are usually normal for modern drives.",
    "Current_Pending_Sector":
        "CRITICAL: Sectors waiting to be tested and possibly remapped. "
        "These are sectors where a read error occurred but the drive hasn't "
        "decided yet if they're bad. Non-zero = data at risk.",
    "Offline_Uncorrectable":
        "CRITICAL: Sectors that could not be read during offline testing. "
        "These sectors contain unrecoverable data. Non-zero = data loss risk.",
    "UDMA_CRC_Error_Count":
        "CRC errors during data transfer over the SATA cable. "
        "Usually caused by a damaged or poorly connected SATA cable. "
        "Try replacing the cable before blaming the drive.",
    "Multi_Zone_Error_Rate":
        "Rate of errors in multi-zone writes. "
        "Vendor-specific metric; not universally meaningful.",
    "Reallocated_Event_Count":
        "Number of remap (sector replacement) operations performed. "
        "Each event indicates the drive replaced a bad sector with a spare.",
    "Wear_Leveling_Count":
        "SSD wear indicator showing how evenly flash cells are used. "
        "100 = new, decreases over time. Below 10 = drive is near end of life.",
    "Used_Rsvd_Blk_Cnt_Tot":
        "Total number of reserved (spare) blocks that have been used. "
        "Higher values mean more bad blocks were replaced.",
    "Unused_Rsvd_Blk_Cnt_Tot":
        "Remaining spare blocks available. "
        "When this reaches 0, the drive cannot replace any more bad sectors.",
    "Program_Fail_Cnt_Total":
        "Total number of flash programming failures (SSD). "
        "Non-zero on a new drive is concerning.",
    "Erase_Fail_Count_Total":
        "Total number of flash erase failures (SSD). "
        "Non-zero on a new drive is concerning.",
    "Total_LBAs_Written":
        "Total amount of data written to the drive over its lifetime.",
    "Total_LBAs_Read":
        "Total amount of data read from the drive over its lifetime.",
    "Percent_Lifetime_Remain":
        "Estimated remaining lifetime of the drive as a percentage.",
}

# ---------------------------------------------------------------------------
# Health Verdict
# ---------------------------------------------------------------------------

VERDICT_GREEN = "green"
VERDICT_YELLOW = "yellow"
VERDICT_RED = "red"

CPU_TEMP_YELLOW = 60.0
CPU_TEMP_RED = 80.0
GPU_TEMP_YELLOW = 55.0
GPU_TEMP_RED = 75.0
DISK_HEALTH_YELLOW = 90
DISK_HEALTH_RED = 70
DISK_POH_YELLOW = 20_000
DISK_POH_RED = 40_000
BATTERY_WEAR_YELLOW = 20
BATTERY_WEAR_RED = 40


def _temp_verdict(temp, yellow_t, red_t) -> str:
    if not isinstance(temp, (int, float)):
        return VERDICT_GREEN
    if temp > red_t:
        return VERDICT_RED
    if temp > yellow_t:
        return VERDICT_YELLOW
    return VERDICT_GREEN


def _worst(*verdicts) -> str:
    if VERDICT_RED in verdicts:
        return VERDICT_RED
    if VERDICT_YELLOW in verdicts:
        return VERDICT_YELLOW
    return VERDICT_GREEN


def analyze_health(data: dict,
                   smart_details: list[dict] | None = None,
                   battery: dict | None = None) -> dict:
    """Compute per-component health verdict and overall grade."""
    components: dict[str, dict | None] = {}

    # --- CPU ---
    cpu = data.get("cpu", {})
    cpu_temps = cpu.get("temperatures", [])
    cpu_max = max(
        (t["value"] for t in cpu_temps if isinstance(t.get("value"), (int, float))),
        default=0,
    )
    cpu_v = _temp_verdict(cpu_max, CPU_TEMP_YELLOW, CPU_TEMP_RED)
    components["cpu"] = {
        "verdict": cpu_v,
        "details": f"Idle temp: {cpu_max}\u00b0C" if cpu_max else "No temp data",
    }

    # --- GPU ---
    gpus = data.get("gpu", [])
    gpu_v = VERDICT_GREEN
    gpu_parts = []
    for gpu in gpus:
        t = gpu.get("temperature", "N/A")
        v = _temp_verdict(t, GPU_TEMP_YELLOW, GPU_TEMP_RED)
        gpu_v = _worst(gpu_v, v)
        if isinstance(t, (int, float)):
            gpu_parts.append(f"{gpu['name']}: {t}\u00b0C")
    components["gpu"] = {
        "verdict": gpu_v,
        "details": "; ".join(gpu_parts) if gpu_parts else "No temp data",
    }

    # --- RAM ---
    ram = data.get("ram", {})
    sticks = ram.get("sticks", [])
    ram_v = VERDICT_GREEN
    ram_parts: list[str] = []
    if len(sticks) >= 2:
        speeds = {s["speed_mhz"] for s in sticks if s["speed_mhz"] != "N/A"}
        caps = {s["capacity_gb"] for s in sticks if s["capacity_gb"] != "N/A"}
        if len(speeds) > 1:
            ram_v = VERDICT_YELLOW
            ram_parts.append(f"Mismatched speeds: {speeds}")
        if len(caps) > 1:
            ram_v = VERDICT_YELLOW
            ram_parts.append(f"Mismatched capacities: {caps}")
        ch = ram.get("channel_mode", "")
        if "Single" in ch:
            ram_v = _worst(ram_v, VERDICT_YELLOW)
            ram_parts.append("Single Channel with multiple sticks")
    components["ram"] = {
        "verdict": ram_v,
        "details": "; ".join(ram_parts) if ram_parts else "OK",
    }

    # --- Disk ---
    disks = data.get("disks", [])
    disk_v = VERDICT_GREEN
    disk_parts: list[str] = []
    for i, disk in enumerate(disks):
        d_v = VERDICT_GREEN
        health = disk.get("health_pct", "N/A")
        if isinstance(health, str) and health == "FAILING":
            d_v = VERDICT_RED
            disk_parts.append(f"Disk {i+1}: FAILING")
        elif isinstance(health, (int, float)):
            if health < DISK_HEALTH_RED:
                d_v = VERDICT_RED
                disk_parts.append(f"Disk {i+1}: health {health}%")
            elif health < DISK_HEALTH_YELLOW:
                d_v = _worst(d_v, VERDICT_YELLOW)
                disk_parts.append(f"Disk {i+1}: health {health}%")

        poh = disk.get("power_on_hours", "N/A")
        if isinstance(poh, (int, float)):
            if poh > DISK_POH_RED:
                d_v = _worst(d_v, VERDICT_RED)
                disk_parts.append(f"Disk {i+1}: {poh}h power-on")
            elif poh > DISK_POH_YELLOW:
                d_v = _worst(d_v, VERDICT_YELLOW)
                disk_parts.append(f"Disk {i+1}: {poh}h power-on")

        if smart_details and i < len(smart_details):  # type: ignore
            sd = smart_details[i]  # type: ignore
            ca = sd.get("critical_attrs", {})
            if ca.get("reallocated_sectors", 0) > 0:
                d_v = _worst(d_v, VERDICT_RED)
                disk_parts.append(f"Disk {i+1}: {ca['reallocated_sectors']} reallocated sectors")
            if ca.get("pending_sectors", 0) > 0:
                d_v = _worst(d_v, VERDICT_RED)
                disk_parts.append(f"Disk {i+1}: {ca['pending_sectors']} pending sectors")
            if ca.get("crc_errors", 0) > 0:
                d_v = _worst(d_v, VERDICT_YELLOW)
                disk_parts.append(f"Disk {i+1}: {ca['crc_errors']} CRC errors")

        disk_v = _worst(disk_v, d_v)

    components["disk"] = {
        "verdict": disk_v,
        "details": "; ".join(disk_parts) if disk_parts else "OK",
    }

    # --- Battery ---
    if battery and battery.get("detected"):  # type: ignore
        wear = battery.get("wear_pct", 0)
        if wear > BATTERY_WEAR_RED:
            bat_v = VERDICT_RED
        elif wear > BATTERY_WEAR_YELLOW:
            bat_v = VERDICT_YELLOW
        else:
            bat_v = VERDICT_GREEN
        components["battery"] = {
            "verdict": bat_v,
            "details": f"Wear: {wear}%, Cycles: {battery.get('cycle_count', 'N/A')}",
        }
    else:
        components["battery"] = None

    # --- Overall ---
    comp_verdicts = [c["verdict"] for c in components.values() if c]
    overall = _worst(*comp_verdicts) if comp_verdicts else VERDICT_GREEN

    reds = sum(1 for v in comp_verdicts if v == VERDICT_RED)
    yellows = sum(1 for v in comp_verdicts if v == VERDICT_YELLOW)
    if reds >= 2:
        grade = "F"
    elif reds == 1:
        grade = "D"
    elif yellows >= 3:
        grade = "C"
    elif yellows >= 1:
        grade = "B"
    else:
        grade = "A"

    return {"overall": overall, "grade": grade, "components": components}


# ---------------------------------------------------------------------------
# Deep S.M.A.R.T.
# ---------------------------------------------------------------------------

_CRITICAL_ATA_ATTRS = {
    "Reallocated_Sector_Ct": "reallocated_sectors",
    "Current_Pending_Sector": "pending_sectors",
    "Offline_Uncorrectable": "uncorrectable_sectors",
    "UDMA_CRC_Error_Count": "crc_errors",
    "Reallocated_Event_Count": "reallocated_events",
    "Raw_Read_Error_Rate": "read_error_rate",
    "Spin_Retry_Count": "spin_retry_count",
}


def get_deep_smart() -> list[dict]:
    """Full S.M.A.R.T. attribute dump for every detected disk."""
    if not _find_smartctl():
        return []

    devices = _smartctl_scan()
    results: list[dict] = []

    for device in devices:
        smart = _smartctl_json(device)
        if not smart:
            results.append({"device": device, "model": "N/A", "error": "Cannot read"})
            continue

        info: dict[str, Any] = {
            "device": device,
            "model": smart.get("model_name", "N/A"),
            "protocol": smart.get("device", {}).get("protocol", "Unknown"),
            "smart_passed": smart.get("smart_status", {}).get("passed"),
            "critical_attrs": {},
            "all_attrs": [],
            "nvme_info": None,
        }

        for attr in smart.get("ata_smart_attributes", {}).get("table", []):
            name = attr.get("name", "")
            raw_val = attr.get("raw", {}).get("value", 0)
            info["all_attrs"].append({
                "id": attr.get("id", 0),
                "name": name,
                "value": attr.get("value", 0),
                "worst": attr.get("worst", 0),
                "thresh": attr.get("thresh", 0),
                "raw": raw_val,
            })
            if name in _CRITICAL_ATA_ATTRS:
                info["critical_attrs"][_CRITICAL_ATA_ATTRS[name]] = int(raw_val)

        for key in _CRITICAL_ATA_ATTRS.values():
            info["critical_attrs"].setdefault(key, 0)

        nvme_log = smart.get("nvme_smart_health_information_log", {})
        if nvme_log:
            info["nvme_info"] = {
                "percentage_used": nvme_log.get("percentage_used", 0),
                "media_errors": nvme_log.get("media_errors", 0),
                "unsafe_shutdowns": nvme_log.get("unsafe_shutdowns", 0),
                "power_on_hours": nvme_log.get("power_on_hours", 0),
                "data_read_tb": round(
                    nvme_log.get("data_units_read", 0) * 512 / 1e12, 2),
                "data_written_tb": round(
                    nvme_log.get("data_units_written", 0) * 512 / 1e12, 2),
            }
            info["critical_attrs"]["media_errors"] = nvme_log.get("media_errors", 0)
            info["critical_attrs"]["percentage_used"] = nvme_log.get(
                "percentage_used", 0)

        results.append(info)

    return results


# ---------------------------------------------------------------------------
# Red Flags
# ---------------------------------------------------------------------------

def detect_red_flags(data: dict,
                     smart_details: list[dict] | None = None) -> list[dict]:
    """Detect suspicious patterns; returns [{severity, component, message}]."""
    flags: list[dict] = []

    sticks = data.get("ram", {}).get("sticks", [])
    if len(sticks) >= 2:
        speeds = [s["speed_mhz"] for s in sticks if s["speed_mhz"] != "N/A"]
        if len(set(speeds)) > 1:
            flags.append({
                "severity": "warning", "component": "RAM",
                "message_key": "flag_ram_speed_mismatch",
                "message_args": {"speeds": ", ".join(str(x) for x in speeds)},
            })
        caps = [s["capacity_gb"] for s in sticks if s["capacity_gb"] != "N/A"]
        if len(set(caps)) > 1:
            flags.append({
                "severity": "warning", "component": "RAM",
                "message_key": "flag_ram_cap_mismatch",
                "message_args": {"caps": ", ".join(str(x) for x in caps)},
            })
        mfrs = [s["manufacturer"] for s in sticks
                if s["manufacturer"] not in ("N/A", "")]
        if len(set(mfrs)) > 1:
            flags.append({
                "severity": "warning", "component": "RAM",
                "message_key": "flag_ram_mfr_mismatch",
                "message_args": {"mfrs": ", ".join(mfrs)},
            })

    _SUSPICIOUS_SN = {"", "0", "00000000", "N/A", "Unknown", "None"}
    for i, disk in enumerate(data.get("disks", [])):
        sn = str(disk.get("serial", "")).strip()
        if sn in _SUSPICIOUS_SN:
            flags.append({
                "severity": "warning", "component": "Disk",
                "message_key": "flag_suspicious_serial",
                "message_args": {"part": f"Disk #{i+1}"},
            })
    _missing_ram_sn = sum(
        1 for s in sticks
        if str(s.get("serial", "")).strip() in _SUSPICIOUS_SN
    )
    if _missing_ram_sn > 0:
        flags.append({
            "severity": "info", "component": "RAM",
            "message_key": "flag_ram_no_serial",
            "message_args": {"count": str(_missing_ram_sn)},
        })

    for i, disk in enumerate(data.get("disks", [])):
        health = disk.get("health_pct", "N/A")
        if isinstance(health, str) and health == "FAILING":
            flags.append({
                "severity": "critical", "component": "Disk",
                "message_key": "flag_smart_failing",
                "message_args": {"disk": f"#{i+1} ({disk['model']})"},
            })
        poh = disk.get("power_on_hours", "N/A")
        if isinstance(poh, (int, float)) and poh > DISK_POH_RED:
            flags.append({
                "severity": "warning", "component": "Disk",
                "message_key": "flag_high_poh",
                "message_args": {"disk": f"#{i+1}", "hours": str(poh)},
            })

    if smart_details:
        for i, sd in enumerate(smart_details):  # type: ignore
            if sd.get("error"):
                continue
            ca = sd.get("critical_attrs", {})
            if ca.get("reallocated_sectors", 0) > 0:
                flags.append({
                    "severity": "critical", "component": "Disk",
                    "message_key": "flag_reallocated",
                    "message_args": {"disk": f"#{i+1}",
                                     "count": str(ca["reallocated_sectors"])},
                })
            if ca.get("pending_sectors", 0) > 0:
                flags.append({
                    "severity": "critical", "component": "Disk",
                    "message_key": "flag_pending",
                    "message_args": {"disk": f"#{i+1}",
                                     "count": str(ca["pending_sectors"])},
                })
            if ca.get("uncorrectable_sectors", 0) > 0:
                flags.append({
                    "severity": "critical", "component": "Disk",
                    "message_key": "flag_uncorrectable",
                    "message_args": {"disk": f"#{i+1}",
                                     "count": str(ca["uncorrectable_sectors"])},
                })
            if ca.get("crc_errors", 0) > 0:
                flags.append({
                    "severity": "warning", "component": "Disk",
                    "message_key": "flag_crc",
                    "message_args": {"disk": f"#{i+1}",
                                     "count": str(ca["crc_errors"])},
                })
            if ca.get("spin_retry_count", 0) > 0:
                flags.append({
                    "severity": "warning", "component": "Disk",
                    "message_key": "flag_spin_retry",
                    "message_args": {"disk": f"#{i+1}",
                                     "count": str(ca["spin_retry_count"])},
                })
            nvme = sd.get("nvme_info")
            if nvme:
                if nvme.get("media_errors", 0) > 0:
                    flags.append({
                        "severity": "critical", "component": "Disk",
                        "message_key": "flag_nvme_media",
                        "message_args": {"disk": f"#{i+1}",
                                         "count": str(nvme["media_errors"])},
                    })
                if nvme.get("percentage_used", 0) > 90:
                    flags.append({
                        "severity": "critical", "component": "Disk",
                        "message_key": "flag_nvme_worn",
                        "message_args": {"disk": f"#{i+1}",
                                         "pct": str(nvme["percentage_used"])},
                    })

    import datetime
    mb = data.get("motherboard", {})
    bdate = mb.get("bios_date", "N/A")
    if bdate != "N/A" and len(bdate) >= 4:
        try:
            byear = int(bdate[:4])
            if byear > datetime.datetime.now().year:
                flags.append({
                    "severity": "warning", "component": "Motherboard",
                    "message_key": "flag_bios_future",
                    "message_args": {"date": bdate},
                })
        except ValueError:
            pass

    for i, gpu in enumerate(data.get("gpu", [])):
        vram = gpu.get("vram_gb", "N/A")
        if isinstance(vram, (int, float)) and 0 < vram < 0.5:
            flags.append({
                "severity": "warning", "component": "GPU",
                "message_key": "flag_low_vram",
                "message_args": {"gpu": f"#{i+1}", "vram": str(vram)},
            })

    return flags


# ---------------------------------------------------------------------------
# Battery Health
# ---------------------------------------------------------------------------

_BATTERY_STATUS = {
    1: "Discharging", 2: "AC Power", 3: "Fully Charged", 4: "Low",
    5: "Critical", 6: "Charging", 7: "Charging (High)",
    8: "Charging (Low)", 9: "Charging (Critical)", 10: "Undefined",
    11: "Partially Charged",
}


def get_battery_info() -> dict | None:
    """Return battery health dict or *None* for desktops."""
    pythoncom.CoInitialize()
    try:
        try:
            import wmi as wmi_module  # type: ignore
            w = wmi_module.WMI(namespace="root/cimv2")
            batteries = list(w.Win32_Battery())
            if not batteries:
                return None
        except Exception:
            return None

        bat = batteries[0]
        info: dict[str, Any] = {
            "detected": True,
            "status": _BATTERY_STATUS.get(int(bat.BatteryStatus or 0), "Unknown"),
            "charge_pct": int(bat.EstimatedChargeRemaining or 0),
            "design_capacity": "N/A",
            "full_charge_capacity": "N/A",
            "wear_pct": 0,
            "cycle_count": "N/A",
        }

        try:
            w2 = wmi_module.WMI(namespace="root/WMI")
            for sd in w2.BatteryStaticData():
                info["design_capacity"] = int(sd.DesignedCapacity or 0)
                break
            for fc in w2.BatteryFullChargedCapacity():
                info["full_charge_capacity"] = int(fc.FullChargedCapacity or 0)
                break
            for cc in w2.BatteryCycleCount():
                info["cycle_count"] = int(cc.CycleCount or 0)
                break
            dc = info["design_capacity"]
            fc = info["full_charge_capacity"]
            if isinstance(dc, int) and dc > 0 and isinstance(fc, int):
                info["wear_pct"] = max(round((1 - fc / dc) * 100, 1), 0)  # type: ignore
        except Exception:
            pass

        return info
    finally:
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# CPU Stress Test
# ---------------------------------------------------------------------------

def _cpu_burn(stop: threading.Event):
    x = 0.0
    while not stop.is_set():
        for _ in range(100_000):
            x += math.sin(x) * math.cos(x) + 1.0  # type: ignore
            if stop.is_set():
                break


class CpuStressTest:
    """Runs all-core CPU stress with temperature monitoring."""

    def __init__(self, duration_sec: int,
                 on_tick: Callable[..., Any],
                 on_done: Callable[..., Any]):
        self.duration = duration_sec
        self.on_tick = on_tick
        self.on_done = on_done
        self._stop = threading.Event()
        self._workers: list[threading.Thread] = []

    def start(self):
        n = psutil.cpu_count(logical=True) or 4
        self._stop.clear()
        for _ in range(n):
            t = threading.Thread(target=_cpu_burn, args=(self._stop,), daemon=True)
            t.start()
            self._workers.append(t)
        threading.Thread(target=self._monitor, daemon=True).start()

    def stop(self):
        self._stop.set()

    def _monitor(self):
        pythoncom.CoInitialize()
        try:
            _get_cpu_usage()
            max_temp = 0.0
            min_clock = 999_999.0
            max_clock = 0.0
            start_temps: list[float] = []
            throttled = False
            t0 = time.time()

            time.sleep(1)

            for sec in range(self.duration):
                if self._stop.is_set():
                    break
                sensors = _get_lhm_sensors()
                cpu_t = _find_sensors(sensors, sensor_type="Temperature",
                                      parent_contains="cpu")
                temps = [round(t["Value"], 1) for t in cpu_t
                         if t["Value"] is not None]
                cur_max = max(temps) if temps else 0.0
                if cur_max > max_temp:
                    max_temp = cur_max
                if sec == 0:
                    start_temps = temps[:]  # type: ignore

                clock = 0.0
                try:
                    import wmi as wmi_module  # type: ignore
                    w = wmi_module.WMI(namespace="root/cimv2")
                    perf = w.Win32_PerfFormattedData_Counters_ProcessorInformation(
                        Name="_Total")
                    for p in perf:
                        base = int(p.ProcessorFrequency or 0)
                        ppp = int(p.PercentProcessorPerformance or 0)
                        if base and ppp:
                            clock = round(base * ppp / 100.0)
                            if clock < min_clock:  # type: ignore
                                min_clock = clock
                            if clock > max_clock:  # type: ignore
                                max_clock = clock
                        break
                except Exception:
                    pass

                usage, _ = _get_cpu_usage()
                self.on_tick(sec + 1, temps, clock, usage)

                if sec > 5 and max_clock > 0 and clock < max_clock * 0.85:  # type: ignore
                    throttled = True

                time.sleep(1)

            self._stop.set()
            for w in self._workers:
                w.join(timeout=2)

            self.on_done({
                "duration_actual": min(self.duration, int(time.time() - t0)),
                "max_temp": max_temp,
                "start_temp": max(start_temps) if start_temps else 0,
                "temp_delta": round(  # type: ignore
                    max_temp - (max(start_temps) if start_temps else 0), 1),  # type: ignore
                "min_clock": min_clock if min_clock < 999_999 else 0,  # type: ignore
                "max_clock": max_clock,
                "throttled": throttled,
                "passed": not throttled and max_temp < 100,
            })
        finally:
            pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# Disk Speed Test
# ---------------------------------------------------------------------------

def _win_read_no_cache(filepath: str, block_size: int, total_blocks: int,
                       on_progress: Callable | None) -> float:  # type: ignore
    """Read file bypassing OS cache via FILE_FLAG_NO_BUFFERING. Returns seconds."""
    kernel32 = ctypes.windll.kernel32  # type: ignore
    kernel32.CreateFileW.restype = ctypes.wintypes.HANDLE
    kernel32.CreateFileW.argtypes = [
        ctypes.wintypes.LPCWSTR, ctypes.wintypes.DWORD, ctypes.wintypes.DWORD,
        ctypes.c_void_p, ctypes.wintypes.DWORD, ctypes.wintypes.DWORD,
        ctypes.wintypes.HANDLE,
    ]
    GENERIC_READ = 0x80000000
    OPEN_EXISTING = 3
    FILE_FLAG_NO_BUFFERING = 0x20000000
    FILE_FLAG_SEQUENTIAL_SCAN = 0x08000000

    handle = kernel32.CreateFileW(
        filepath, GENERIC_READ, 1, None,
        OPEN_EXISTING, FILE_FLAG_NO_BUFFERING | FILE_FLAG_SEQUENTIAL_SCAN, None)

    INVALID = ctypes.wintypes.HANDLE(-1).value
    if handle == INVALID or handle is None:
        raise OSError("Cannot open file with FILE_FLAG_NO_BUFFERING")

    try:
        buf = ctypes.create_string_buffer(block_size)
        bytes_read = ctypes.wintypes.DWORD()
        t0 = time.perf_counter()
        i = 0
        while True:
            ok = kernel32.ReadFile(handle, buf, block_size,
                                   ctypes.byref(bytes_read), None)
            if not ok or bytes_read.value == 0:
                break
            i += 1  # type: ignore
            if on_progress and i % 32 == 0:
                on_progress("read", i, total_blocks)  # type: ignore
        return time.perf_counter() - t0
    finally:
        kernel32.CloseHandle(handle)


def run_disk_speed_test(drive_path: str, size_mb: int = 256,
                        on_progress: Callable | None = None,
                        on_done: Callable | None = None):
    def _worker():
        results: dict[str, Any] = {"write_mbps": 0, "read_mbps": 0, "error": None}
        try:
            tmp_dir = drive_path if os.path.isdir(drive_path) else os.path.dirname(
                drive_path)
            if not tmp_dir or not os.path.exists(tmp_dir):
                tmp_dir = tempfile.gettempdir()

            test_file = os.path.join(tmp_dir, "_hwdiag_speedtest.tmp")
            block_size = 1024 * 1024
            block = os.urandom(block_size)
            total = size_mb

            # --- Write test ---
            if on_progress:
                on_progress("write", 0, total)
            t0 = time.perf_counter()
            with open(test_file, "wb") as f:
                for i in range(total):
                    f.write(block)
                    if on_progress and i % 32 == 0:
                        on_progress("write", i + 1, total)  # type: ignore
                f.flush()
                os.fsync(f.fileno())
            results["write_mbps"] = round(size_mb / (time.perf_counter() - t0), 1)  # type: ignore
            if on_progress:
                on_progress("write", total, total)  # type: ignore

            # --- Read test (bypass OS cache) ---
            if on_progress:
                on_progress("read", 0, total)
            try:
                dt = _win_read_no_cache(test_file, block_size, total, on_progress)
            except OSError:
                t0 = time.perf_counter()
                i = 0
                with open(test_file, "rb") as f:
                    while f.read(block_size):
                        i += 1
                        if on_progress and i % 32 == 0:
                            on_progress("read", i, total)  # type: ignore
                dt = time.perf_counter() - t0
            results["read_mbps"] = round(size_mb / dt, 1) if dt > 0 else 0  # type: ignore
            if on_progress:
                on_progress("read", total, total)

            try:
                os.remove(test_file)
            except Exception:
                pass
        except Exception as e:
            results["error"] = str(e)

        if on_done:
            on_done(results)

    threading.Thread(target=_worker, daemon=True).start()


# ---------------------------------------------------------------------------
# RAM Quick Check
# ---------------------------------------------------------------------------

def run_ram_check(size_mb: int = 0,
                  on_progress: Callable | None = None,
                  on_done: Callable | None = None):
    """Test RAM for bit-flip errors.

    *size_mb=0* means auto (25 % of available RAM, clamped 512..4096 MB).
    """
    def _worker():
        results: dict[str, Any] = {
            "tested_mb": 0, "errors": 0, "error": None, "duration_sec": 0}
        try:
            actual_mb = size_mb
            if actual_mb <= 0:
                avail = psutil.virtual_memory().available // (1024 * 1024)
                actual_mb = max(512, min(int(avail * 0.70), 8192))

            patterns = [0xAA, 0x55, 0xFF, 0x00, 0x0F, 0xF0]
            MB = 1024 * 1024
            total_ops = actual_mb * len(patterns) * 2  # alloc + verify
            done_ops = 0
            errors = 0
            t0 = time.perf_counter()

            for pattern in patterns:
                ref = bytes([pattern]) * MB
                buffers: list[bytearray] = []
                allocated = 0
                for _ in range(actual_mb):
                    try:
                        buffers.append(bytearray(ref))
                        allocated += 1
                    except MemoryError:
                        break
                    done_ops += 1  # type: ignore
                    if on_progress and done_ops % 32 == 0:
                        on_progress(done_ops, total_ops)

                for buf in buffers:
                    if bytes(buf) != ref:
                        errors += 1  # type: ignore
                    done_ops += 1
                    if on_progress and done_ops % 32 == 0:
                        on_progress(done_ops, total_ops)  # type: ignore

                buffers.clear()

            results["tested_mb"] = actual_mb
            results["errors"] = errors
            results["duration_sec"] = round(time.perf_counter() - t0, 1)  # type: ignore
            if on_progress:
                on_progress(total_ops, total_ops)  # type: ignore
        except Exception as e:
            results["error"] = str(e)

        if on_done:
            on_done(results)

    threading.Thread(target=_worker, daemon=True).start()


# ---------------------------------------------------------------------------
# CPU Benchmark Lite
# ---------------------------------------------------------------------------

def _load_benchmark_db() -> dict:
    """Load CPU benchmark reference scores from JSON file."""
    try:
        path = get_resource_path("benchmarks.json")
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def _match_cpu_name(cpu_name: str, db: dict) -> tuple[str, dict] | None:
    """Fuzzy-match a CPU name against the benchmark database keys."""
    if not cpu_name or cpu_name == "N/A":
        return None
    name_lower = cpu_name.lower()
    for key in db:
        if key.lower() in name_lower:
            return key, db[key]
    return None


def run_cpu_benchmark(on_progress: Callable | None = None,
                      on_done: Callable | None = None):
    """Quick CPU benchmark (~5s). Measures single-thread and multi-thread math ops."""
    def _worker():
        results: dict[str, Any] = {
            "single_score": 0, "multi_score": 0,
            "ref_name": None, "ref_single": 0, "ref_multi": 0,
            "error": None,
        }
        try:
            iters = 2_000_000
            n_cores = psutil.cpu_count(logical=True) or 4

            # Single-thread
            if on_progress:
                on_progress("single", 0)
            t0 = time.perf_counter()
            x = 1.0
            for i in range(iters):
                x = math.sin(x) * math.cos(x) + math.sqrt(abs(x) + 1.0)  # type: ignore
            single_time = time.perf_counter() - t0
            results["single_score"] = round(iters / single_time / 1000)
            if on_progress:
                on_progress("single", 100)

            # Multi-thread
            if on_progress:
                on_progress("multi", 0)
            barrier = threading.Barrier(n_cores + 1)
            scores = [0] * n_cores

            def _bench_thread(idx):
                barrier.wait()
                t = time.perf_counter()
                v = 1.0
                for _ in range(iters):
                    v = math.sin(v) * math.cos(v) + math.sqrt(abs(v) + 1.0)  # type: ignore
                scores[idx] = round(iters / (time.perf_counter() - t) / 1000)

            threads = []
            for i in range(n_cores):
                t = threading.Thread(target=_bench_thread, args=(i,), daemon=True)
                t.start()
                threads.append(t)
            barrier.wait()
            for t in threads:
                t.join()
            results["multi_score"] = sum(scores)
            if on_progress:
                on_progress("multi", 100)

        except Exception as e:
            results["error"] = str(e)

        if on_done:
            on_done(results)

    threading.Thread(target=_worker, daemon=True).start()


# ---------------------------------------------------------------------------
# System File Integrity (SFC / DISM)
# ---------------------------------------------------------------------------

def _decode_console(raw: bytes) -> str:
    """Decode bytes from a Windows console subprocess (cp866 -> cp1251 fallback)."""
    for enc in ("cp866", "cp1251", "utf-8"):
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("utf-8", errors="replace")


def run_sfc_scan(on_line: Callable | None = None,
                 on_done: Callable | None = None):
    """Run sfc /scannow in background. Requires admin."""
    def _worker():
        result = {"output": "", "error": None, "violations_found": False}
        try:
            proc = subprocess.Popen(
                ["sfc", "/scannow"],
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                creationflags=subprocess.CREATE_NO_WINDOW,  # type: ignore
            )
            lines = []
            for raw_line in proc.stdout:  # type: ignore
                line = _decode_console(raw_line).rstrip()  # type: ignore
                if line:
                    lines.append(line)
                    if on_line:
                        on_line(line)
            proc.wait()
            output = "\n".join(lines)
            result["output"] = output
            low = output.lower()
            if "found integrity violations" in low or "нарушения целостности" in low:
                result["violations_found"] = True
        except Exception as e:
            result["error"] = str(e)
        if on_done:
            on_done(result)

    threading.Thread(target=_worker, daemon=True).start()


def run_dism_scan(on_line: Callable | None = None,
                  on_done: Callable | None = None):
    """Run DISM /Online /Cleanup-Image /ScanHealth in background."""
    def _worker():
        result = {"output": "", "error": None, "repairable": False}
        try:
            proc = subprocess.Popen(
                ["DISM", "/Online", "/Cleanup-Image", "/ScanHealth"],
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                creationflags=subprocess.CREATE_NO_WINDOW,  # type: ignore
            )
            lines = []
            for raw_line in proc.stdout:  # type: ignore
                line = _decode_console(raw_line).rstrip()  # type: ignore
                if line:
                    lines.append(line)
                    if on_line:
                        on_line(line)
            proc.wait()
            output = "\n".join(lines)
            result["output"] = output
            low = output.lower()
            if "repairable" in low or "component store corruption" in low:
                result["repairable"] = True
        except Exception as e:
            result["error"] = str(e)
        if on_done:
            on_done(result)

    threading.Thread(target=_worker, daemon=True).start()


# ---------------------------------------------------------------------------
# Event Viewer — Recent Critical Errors
# ---------------------------------------------------------------------------

def get_recent_critical_events(count: int = 10) -> list[dict]:  # type: ignore
    """Fetch recent Error/Critical events from System & Application logs via WMI."""
    pythoncom.CoInitialize()
    try:
        import wmi as wmi_module  # type: ignore
        events: list[dict] = []
        w = wmi_module.WMI()
        for log_name in ("System", "Application"):
            try:
                query = (
                    f"SELECT * FROM Win32_NTLogEvent "
                    f"WHERE Logfile='{log_name}' AND "
                    f"(EventType=1 OR EventType=2) "
                )
                for ev in w.query(query):
                    ts = ev.TimeGenerated or ""
                    time_str = ""
                    if ts and len(ts) >= 14:
                        try:
                            time_str = f"{ts[:4]}-{ts[4:6]}-{ts[6:8]} {ts[8:10]}:{ts[10:12]}"
                        except Exception:
                            time_str = ts[:14]
                    msg = (ev.Message or "")[:300]
                    events.append({
                        "source": ev.SourceName or "Unknown",
                        "event_id": ev.EventCode or 0,
                        "type": "Critical" if ev.EventType == 1 else "Error",
                        "log": log_name,
                        "time": time_str,
                        "message": msg,
                    })
                    if len(events) >= count * 2:
                        break
            except Exception:
                continue
        events.sort(key=lambda e: e.get("time", ""), reverse=True)
        return events[:count]  # type: ignore
    except Exception:
        return []
    finally:
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# PDF Report
# ---------------------------------------------------------------------------

def generate_pdf(filepath: str, data: dict, *,
                 verdict: dict | None = None,
                 red_flags: list[dict] | None = None,
                 smart_details: list[dict] | None = None,
                 battery: dict | None = None,
                 stress_results: dict | None = None,
                 disk_speed: dict | None = None,
                 ram_check: dict | None = None,
                 lang: str = "en",
                 t_fn: Callable | None = None) -> bool:
    """Generate PDF report. Returns True on success."""
    try:
        from fpdf import FPDF  # type: ignore
    except ImportError:
        return False

    class PDF(FPDF):
        def header(self):
            self.set_font("Helvetica", "B", 16)
            self.cell(0, 10, "Hardware Diagnostics Report", align="C", new_x="LMARGIN", new_y="NEXT")
            self.set_font("Helvetica", "", 9)
            import datetime
            self.cell(0, 6,
                      datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                      align="C", new_x="LMARGIN", new_y="NEXT")
            self.ln(4)

        def footer(self):
            self.set_y(-15)
            self.set_font("Helvetica", "I", 8)
            self.cell(0, 10, f"Page {self.page_no()}/{{nb}}", align="C")

    pdf = PDF()
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    def _section(title: str):
        pdf.set_font("Helvetica", "B", 13)
        pdf.set_fill_color(230, 230, 230)
        pdf.cell(0, 8, f"  {title}", fill=True, new_x="LMARGIN", new_y="NEXT")
        pdf.ln(2)

    def _kv(key: str, val: str):
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(60, 6, key, new_x="RIGHT")
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(0, 6, val, new_x="LMARGIN", new_y="NEXT")

    # --- Verdict ---
    if verdict:
        _section("Health Verdict")
        grade = verdict.get("grade", "?")
        overall = verdict.get("overall", "green")
        color_map = {"green": (67, 160, 71), "yellow": (255, 167, 38),
                     "red": (229, 57, 53)}
        r, g, b = color_map.get(overall, (0, 0, 0))
        pdf.set_font("Helvetica", "B", 24)
        pdf.set_text_color(r, g, b)
        label = {"green": "GOOD", "yellow": "CAUTION", "red": "PROBLEMS"}.get(
            overall, "?")
        pdf.cell(0, 14, f"Grade: {grade} - {label}", align="C",
                 new_x="LMARGIN", new_y="NEXT")
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Helvetica", "", 10)
        for comp_name, comp in verdict.get("components", {}).items():
            if comp is None:
                continue
            v = comp["verdict"]  # type: ignore
            cr, cg, cb = color_map.get(v, (0, 0, 0))
            pdf.set_text_color(cr, cg, cb)
            sym = {"green": "OK", "yellow": "CAUTION", "red": "PROBLEM"}.get(v, "?")
            pdf.cell(0, 6,
                     f"  {comp_name.upper()}: {sym} - {comp['details']}",  # type: ignore
                     new_x="LMARGIN", new_y="NEXT")
        pdf.set_text_color(0, 0, 0)
        pdf.ln(4)

    # --- Red Flags ---
    if red_flags:
        _section(f"Red Flags ({len(red_flags)})")
        for flag in red_flags:
            sev = flag["severity"].upper()
            comp = flag["component"]
            args = flag.get("message_args", {})
            msg_key = flag.get("message_key", "")
            msg = f"[{sev}] {comp}: {msg_key} {args}"
            if sev == "CRITICAL":
                pdf.set_text_color(229, 57, 53)
            else:
                pdf.set_text_color(255, 167, 38)
            pdf.set_font("Helvetica", "", 9)
            pdf.cell(0, 5, f"  {msg}", new_x="LMARGIN", new_y="NEXT")
        pdf.set_text_color(0, 0, 0)
        pdf.ln(4)

    # --- CPU ---
    cpu = data.get("cpu", {})
    _section("CPU")
    _kv("Name:", cpu.get("name", "N/A"))
    _kv("Cores / Threads:", f"{cpu.get('cores', 'N/A')} / {cpu.get('threads', 'N/A')}")
    _kv("Base Clock:", f"{cpu.get('base_clock_mhz', 'N/A')} MHz")
    _kv("Current Clock:", f"{cpu.get('current_clock_mhz', 'N/A')} MHz")
    for t in cpu.get("temperatures", []):
        _kv(f"  {t['name']}:", f"{t['value']}\u00b0C")
    pdf.ln(2)

    # --- GPU ---
    for i, gpu in enumerate(data.get("gpu", [])):
        _section(f"GPU #{i+1}")
        _kv("Name:", gpu.get("name", "N/A"))
        _kv("Vendor:", gpu.get("vendor", "N/A"))
        _kv("VRAM:", f"{gpu.get('vram_gb', 'N/A')} GB")
        _kv("Temperature:", f"{gpu.get('temperature', 'N/A')}\u00b0C")
        pdf.ln(2)

    # --- RAM ---
    ram = data.get("ram", {})
    _section("RAM")
    _kv("Total:", f"{ram.get('total_gb', 'N/A')} GB")
    _kv("Slots:", f"{ram.get('used_slots', 0)} / {ram.get('total_slots', 'N/A')}")
    _kv("Channel Mode:", ram.get("channel_mode", "N/A"))
    for i, s in enumerate(ram.get("sticks", [])):
        _kv(f"  DIMM #{i+1}:",
            f"{s.get('manufacturer','N/A')} | "
            f"{s.get('capacity_gb','N/A')} GB | "
            f"{s.get('speed_mhz','N/A')} MHz")
    pdf.ln(2)

    # --- Disks ---
    for i, d in enumerate(data.get("disks", [])):
        _section(f"Disk #{i+1}")
        _kv("Model:", d.get("model", "N/A"))
        _kv("Serial:", d.get("serial", "N/A"))
        _kv("Size:", f"{d.get('size_gb', 'N/A')} GB")
        _kv("Type:", d.get("type", "N/A"))
        _kv("Temperature:", f"{d.get('temperature', 'N/A')}\u00b0C")
        _kv("Power-on Hours:", f"{d.get('power_on_hours', 'N/A')}")
        hp = d.get("health_pct", "N/A")
        _kv("Health:", f"{hp}%" if isinstance(hp, (int, float)) else str(hp))
        pdf.ln(2)

    # --- S.M.A.R.T. Details ---
    if smart_details:
        for i, sd in enumerate(smart_details):  # type: ignore
            if sd.get("error"):
                continue
            attrs = sd.get("all_attrs", [])
            if not attrs:
                continue
            _section(f"S.M.A.R.T. - {sd.get('model', f'Disk #{i+1}')}")
            pdf.set_font("Helvetica", "B", 8)
            pdf.cell(15, 5, "ID")
            pdf.cell(55, 5, "Attribute")
            pdf.cell(20, 5, "Value")
            pdf.cell(20, 5, "Worst")
            pdf.cell(20, 5, "Thresh")
            pdf.cell(0, 5, "Raw", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font("Helvetica", "", 8)
            for a in attrs:
                name = a.get("name", "")
                is_critical = name in _CRITICAL_ATA_ATTRS and a.get("raw", 0) > 0
                if is_critical:
                    pdf.set_text_color(229, 57, 53)
                pdf.cell(15, 4, str(a.get("id", "")))
                pdf.cell(55, 4, name[:30])
                pdf.cell(20, 4, str(a.get("value", "")))
                pdf.cell(20, 4, str(a.get("worst", "")))
                pdf.cell(20, 4, str(a.get("thresh", "")))
                pdf.cell(0, 4, str(a.get("raw", "")),
                         new_x="LMARGIN", new_y="NEXT")
                if is_critical:
                    pdf.set_text_color(0, 0, 0)
            pdf.ln(2)

    # --- Battery ---
    if battery and battery.get("detected"):
        _section("Battery")
        _kv("Status:", battery.get("status", "N/A"))
        _kv("Charge:", f"{battery.get('charge_pct', 'N/A')}%")
        _kv("Design Capacity:", f"{battery.get('design_capacity', 'N/A')} mWh")
        _kv("Full Charge:", f"{battery.get('full_charge_capacity', 'N/A')} mWh")
        _kv("Wear:", f"{battery.get('wear_pct', 0)}%")
        _kv("Cycles:", str(battery.get("cycle_count", "N/A")))
        pdf.ln(2)

    # --- Stress Test Results ---
    if stress_results:
        _section("CPU Stress Test")
        passed = stress_results.get("passed", False)
        pdf.set_text_color(67, 160, 71) if passed else pdf.set_text_color(229, 57, 53)
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 7, "PASSED" if passed else "FAILED",
                 new_x="LMARGIN", new_y="NEXT")
        pdf.set_text_color(0, 0, 0)
        _kv("Duration:", f"{stress_results.get('duration_actual', 0)}s")
        _kv("Max Temp:", f"{stress_results.get('max_temp', 0)}\u00b0C")
        _kv("Temp Delta:", f"+{stress_results.get('temp_delta', 0)}\u00b0C")
        _kv("Throttled:", "Yes" if stress_results.get("throttled") else "No")
        pdf.ln(2)

    if disk_speed:
        _section("Disk Speed Test")
        _kv("Write:", f"{disk_speed.get('write_mbps', 0)} MB/s")
        _kv("Read:", f"{disk_speed.get('read_mbps', 0)} MB/s")
        pdf.ln(2)

    if ram_check:
        _section("RAM Check")
        _kv("Tested:", f"{ram_check.get('tested_mb', 0)} MB")
        _kv("Errors:", str(ram_check.get("errors", 0)))
        _kv("Duration:", f"{ram_check.get('duration_sec', 0)}s")
        pdf.ln(2)

    # --- Motherboard ---
    mb = data.get("motherboard", {})
    _section("Motherboard")
    _kv("Manufacturer:", mb.get("manufacturer", "N/A"))
    _kv("Model:", mb.get("model", "N/A"))
    _kv("BIOS Version:", mb.get("bios_version", "N/A"))
    _kv("BIOS Date:", mb.get("bios_date", "N/A"))

    try:
        pdf.output(filepath)
        return True
    except Exception:
        return False
