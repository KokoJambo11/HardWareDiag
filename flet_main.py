import sys
import threading
import time
from datetime import datetime, timedelta
from collections import deque
import flet as ft
import flet.canvas as cv

from hardware import (
    collect_all, collect_dynamic, get_disk_info, is_admin,
    check_lhm_available, check_smartctl_available,
    get_autostart, set_autostart,
)
from diagnostics import (
    analyze_health, detect_red_flags, get_deep_smart, get_battery_info,
    CpuStressTest, run_disk_speed_test, run_ram_check, generate_pdf,
    run_cpu_benchmark, run_sfc_scan, run_dism_scan, get_recent_critical_events,
    VERDICT_GREEN, VERDICT_YELLOW, VERDICT_RED,
)
from translations import T, LANG_NAMES, SMART_DESC, SMART_COL_HELP

MAIN_BG_L = "#f5f5f5"
MAIN_BG_D = "#121212"
CARD_BG_L = "#ffffff"
CARD_BG_D = "#1e1e1e"
ACCENT_L = "#0055ff"
ACCENT_D = "#00d4ff"
ALERT = "#E53935"
WARN = "#FFA726"
GREEN = "#43A047"

def main(page: ft.Page):
    page.title = "Hardware Diagnostics"
    page.window.width = 1100
    page.window.height = 750
    page.theme_mode = ft.ThemeMode.DARK
    page.bgcolor = MAIN_BG_D
    page.scroll = ft.ScrollMode.ADAPTIVE
    page.padding = 20

    lang = "en"
    theme = "dark"

    def t(key: str) -> str:
        return T.get(lang, T["en"]).get(key, T["en"].get(key, key))

    def smart_desc(attr_name: str) -> str:
        """Look up translated S.M.A.R.T. attribute description with EN fallback."""
        lang_d = SMART_DESC.get(lang, SMART_DESC.get("en", {}))
        en_d = SMART_DESC.get("en", {})
        return lang_d.get(attr_name, en_d.get(attr_name, lang_d.get("_generic", en_d.get("_generic", ""))))

    def smart_col_help(col_key: str) -> str:
        """Tooltip for S.M.A.R.T. column header (id, name, value, worst, thresh, raw)."""
        lang_d = SMART_COL_HELP.get(lang, SMART_COL_HELP.get("en", {}))
        return lang_d.get(col_key, SMART_COL_HELP.get("en", {}).get(col_key, ""))

    system_data = {}
    smart_data = []
    battery_data = None
    dyn_labels = []
    cpu_history = deque(maxlen=60)
    gpu_histories = []  # list of deques, one per GPU
    ram_history = deque(maxlen=60)
    
    title_text = ft.Text(t("title"), size=24, weight=ft.FontWeight.BOLD, color=ACCENT_D)
    admin_status = ft.Text(t("admin") if is_admin() else t("no_admin"), color=ft.Colors.GREEN if is_admin() else WARN, size=12)

    def change_lang(e):
        nonlocal lang
        lang = e.control.value
        title_text.value = t("title")
        admin_status.value = t("admin") if is_admin() else t("no_admin")
        page.update()
        build_ui()

    lang_dropdown = ft.Dropdown(
        width=100,
        options=[ft.dropdown.Option(k) for k in LANG_NAMES.keys()],
        value="en",
        on_select=change_lang,
        dense=True
    )

    def toggle_theme(e):
        nonlocal theme
        theme = "light" if theme == "dark" else "dark"
        page.theme_mode = ft.ThemeMode.LIGHT if theme == "light" else ft.ThemeMode.DARK
        page.bgcolor = MAIN_BG_L if theme == "light" else MAIN_BG_D
        title_text.color = ACCENT_L if theme == "light" else ACCENT_D
        e.control.icon = ft.Icons.DARK_MODE if theme == "light" else ft.Icons.LIGHT_MODE
        page.update()
        build_ui()

    theme_btn = ft.IconButton(icon=ft.Icons.LIGHT_MODE, on_click=toggle_theme)

    header_row = ft.Row([
        title_text, ft.Container(expand=True), admin_status, lang_dropdown, theme_btn
    ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)

    dashboard_col = ft.Column(spacing=15, expand=True)

    def make_row(k, v, color=None):
        return ft.Row([
            ft.Text(t(k), width=180, size=13, color=ft.Colors.GREY),
            ft.Text(str(v), size=13, weight=ft.FontWeight.W_500, color=color)
        ])

    def create_progress_bar():
        return ft.ProgressBar(value=0, color=ACCENT_D if theme == "dark" else ACCENT_L, bgcolor=CARD_BG_D if theme == "dark" else CARD_BG_L, height=4)

    CHART_W, CHART_H = 250, 50
    CHART_COLORS = {"cpu": "#2196F3", "gpu": "#2196F3", "ram": "#9C27B0", "disk": "#4CAF50"}
    GRID_LIGHT, GRID_DARK = "#E0E0E0", "#333333"

    def make_chart_path_elements(history, width=CHART_W, height=CHART_H):
        """Build (fill_elements, line_elements) from history (list of 0-100 values)."""
        if not history:
            return [], []
        fill_el = []
        line_el = []
        n = len(history)
        for i, v in enumerate(history):
            x = (i / (n - 1)) * width if n > 1 else 0
            y = height - (float(v) / 100.0) * height
            if not line_el:
                fill_el.append(cv.Path.MoveTo(x, height))
                line_el.append(cv.Path.MoveTo(x, y))
            fill_el.append(cv.Path.LineTo(x, y))
            line_el.append(cv.Path.LineTo(x, y))
        fill_el.append(cv.Path.LineTo(width, height))
        fill_el.append(cv.Path.Close())
        return fill_el, line_el

    def create_usage_chart(history, chart_type="cpu"):
        """Create a line chart with grid and filled area. Returns (container, fill_path, line_path, canvas)."""
        color = CHART_COLORS.get(chart_type, CHART_COLORS["cpu"])
        grid_color = GRID_LIGHT if theme == "light" else GRID_DARK
        fill_el, line_el = make_chart_path_elements(history)
        grid_shapes = []
        for pct in (0.25, 0.5, 0.75):
            y = CHART_H * (1 - pct)
            grid_shapes.append(cv.Line(0, y, CHART_W, y, paint=ft.Paint(stroke_width=1, color=grid_color)))
        for pct in (0.25, 0.5, 0.75):
            x = CHART_W * pct
            grid_shapes.append(cv.Line(x, 0, x, CHART_H, paint=ft.Paint(stroke_width=1, color=grid_color)))
        fill_path = cv.Path(
            elements=fill_el,
            paint=ft.Paint(style=ft.PaintingStyle.FILL, color=color + "4D"),
        )
        line_path = cv.Path(
            elements=line_el,
            paint=ft.Paint(stroke_width=2, style=ft.PaintingStyle.STROKE, color=color),
        )
        can = cv.Canvas(width=CHART_W, height=CHART_H, shapes=grid_shapes + [fill_path, line_path])
        cont = ft.Container(can, padding=ft.padding.only(left=180, top=5, bottom=5))
        return cont, fill_path, line_path, can

    def update_chart(bar, new_val_pct):
        # new_val_pct is 0-100
        bar.value = new_val_pct / 100.0
        if bar.page:
            bar.update()

    def update_usage_chart(fill_path, line_path, canvas, history, new_val_pct):
        history.append(new_val_pct)
        fill_el, line_el = make_chart_path_elements(history)
        fill_path.elements = fill_el
        line_path.elements = line_el
        if canvas.page:
            canvas.update()

    console_output = ft.TextField(
        multiline=True,
        read_only=True,
        text_size=12,
        min_lines=10,
        max_lines=15,
        bgcolor=ft.Colors.BLACK,
        color=ft.Colors.GREEN_ACCENT,
        border_color=ft.Colors.TRANSPARENT,
    )

    def log_msg(msg):
        timestamp = datetime.now().strftime("%H:%M:%S")
        console_output.value = f"[{timestamp}] {msg}\n" + (console_output.value or "")
        if console_output.page:
            console_output.update()

    def build_report_text() -> str:
        """Build plain-text report for clipboard."""
        if not system_data:
            return ""
        data = system_data
        v = analyze_health(system_data, smart_data, battery_data)
        flags = detect_red_flags(system_data, smart_data)
        lines: list[str] = []
        sep = "=" * 50
        status_map = {VERDICT_GREEN: "OK", VERDICT_YELLOW: "CAUTION", VERDICT_RED: "PROBLEM"}
        lines.append(t("title"))
        lines.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        lines.append("")
        overall = v.get("overall", VERDICT_GREEN)
        grade = v.get("grade", "?")
        lines += [sep, f"{t('overall_verdict')}: {grade} — {status_map.get(overall, '?')}", sep]
        for comp_name, comp_data in v.get("components", {}).items():
            if comp_data is None:
                continue
            cv = comp_data["verdict"]
            sym = status_map.get(cv, "?")
            details = comp_data.get("details", "")
            line = f"  {comp_name.upper()}: {sym}"
            if details and details != "OK":
                line += f"  — {details}"
            lines.append(line)
        lines.append("")
        if flags:
            lines += [sep, f"{t('red_flags')} ({len(flags)})", sep]
            for flag in flags:
                sev = flag["severity"].upper()
                comp = flag["component"]
                msg = t(flag["message_key"]).format(**flag.get("message_args", {}))
                lines.append(f"  [{sev}] {comp}: {msg}")
            lines.append("")
        cpu = data.get("cpu", {})
        lines += [sep, t("cpu"), sep,
                  f"{t('name')}:          {cpu.get('name', 'N/A')}",
                  f"Cores/Threads: {cpu.get('cores', 0)} / {cpu.get('threads', 0)}",
                  f"{t('base_clock')}:    {cpu.get('base_clock_mhz', 0)} MHz",
                  f"{t('current_clock')}: {cpu.get('current_clock_mhz', 0)} MHz"]
        for i, gpu in enumerate(data.get("gpu", [])):
            lines += [sep, f"{t('gpu')} #{i+1}", sep,
                      f"{t('name')}:          {gpu.get('name', 'N/A')}",
                      f"{t('vendor')}:        {gpu.get('vendor', 'N/A')}",
                      f"{t('vram')}:          {gpu.get('vram_gb', 'N/A')} GB",
                      f"{t('driver_version')}: {gpu.get('driver_version', 'N/A')}"]
        ram = data.get("ram", {})
        lines += [sep, t("ram"), sep,
                  f"{t('total_available')}: {ram.get('total_gb', 0)} GB",
                  f"{t('slots')}:           {ram.get('used_slots', 0)} / {ram.get('total_slots', 0)}",
                  f"{t('channel_mode')}:    {ram.get('channel_mode', 'N/A')}"]
        for i, d in enumerate(data.get("disks", [])):
            sz = f"{d.get('size_gb', 'N/A')} GB" if d.get("size_gb") not in (None, "N/A") else "N/A"
            temp = d.get("temperature", "N/A")
            temp_str = f"{temp} °C" if temp not in (None, "N/A") else "N/A"
            poh = d.get("power_on_hours", "N/A")
            poh_str = f"{poh} h" if isinstance(poh, (int, float)) else str(poh)
            lines += [sep, f"{t('disk')} #{i+1}", sep,
                      f"{t('model')}:         {d.get('model', 'N/A')}",
                      f"{t('size')}:          {sz}",
                      f"{t('type')}:          {d.get('type', 'N/A')}",
                      f"{t('temperature')}:   {temp_str}",
                      f"{t('power_on_hours')}: {poh_str}"]
        if battery_data and battery_data.get("detected"):
            bat = battery_data
            lines += [sep, t("battery"), sep,
                      f"{t('battery_status')}: {bat.get('status', 'N/A')}",
                      f"{t('battery_charge')}: {bat.get('charge_pct', 0)} %",
                      f"{t('wear_level')}: {bat.get('wear_pct', 0)} %"]
        return "\n".join(lines)

    def copy_report(e):
        text = build_report_text()
        if not text:
            return
        async def _do_copy():
            try:
                await page.clipboard.set(text)
            except Exception:
                try:
                    import pyperclip
                    pyperclip.copy(text)
                except Exception:
                    pass
            page.snack_bar = ft.SnackBar(content=ft.Text(t("copied")), open=True)
            page.update()
        try:
            page.run_task(_do_copy)
        except Exception:
            try:
                import pyperclip
                pyperclip.copy(text)
                page.snack_bar = ft.SnackBar(content=ft.Text(t("copied")), open=True)
                page.update()
            except Exception:
                pass

    def build_ui():
        if not system_data:
            return
            
        bg = CARD_BG_D if theme == "dark" else CARD_BG_L
        acc = ACCENT_D if theme == "dark" else ACCENT_L

        dyn_labels.clear()

        # --- CPU ---
        try:
            cpu = system_data.get("cpu", {})
            cpu_clock = ft.Text(f"{cpu.get('current_clock_mhz', 0)} MHz", size=13, weight=ft.FontWeight.BOLD)
            cpu_usage = ft.Text("0 %", size=13, weight=ft.FontWeight.BOLD)
            cpu_temp = ft.Text("-- \u00b0C", size=13, weight=ft.FontWeight.BOLD)
            cpu_chart_cont, cpu_fill_path, cpu_line_path, cpu_chart_canvas = create_usage_chart(cpu_history, "cpu")
            
            dyn_labels.append({"type": "cpu_clock", "label": cpu_clock})
            dyn_labels.append({"type": "cpu_usage", "label": cpu_usage, "fill_path": cpu_fill_path, "line_path": cpu_line_path, "canvas": cpu_chart_canvas, "history": cpu_history})
            dyn_labels.append({"type": "cpu_temp", "label": cpu_temp})

            cpu_card = ft.Container(
                content=ft.Column([
                    ft.Row([ft.Icon(ft.Icons.SPEED, color=acc, size=20), ft.Text(t('cpu'), size=16, weight="bold", color=acc)]),
                    ft.Divider(),
                    make_row("name", cpu.get("name", "N/A")),
                    make_row("cores_threads", f"{cpu.get('cores', 0)} / {cpu.get('threads', 0)}"),
                    make_row("base_clock", f"{cpu.get('base_clock_mhz', 0)} MHz"),
                    ft.Row([ft.Text(t("current_clock"), width=180, size=13, color=ft.Colors.GREY), cpu_clock]),
                    ft.Row([ft.Text(t("cpu_usage"), width=180, size=13, color=ft.Colors.GREY), cpu_usage]),
                    ft.Row([ft.Text(t("temperature"), width=180, size=13, color=ft.Colors.GREY), cpu_temp]),
                    cpu_chart_cont,
                ]),
                bgcolor=bg, border_radius=10, padding=15, expand=1
            )
        except Exception as e:
            log_msg(f"CPU Card error: {e}")
            cpu_card = ft.Container()

        # --- GPU ---
        gpu_cols = []
        while len(gpu_histories) < len(system_data.get("gpu", [])):
            gpu_histories.append(deque(maxlen=60))
        for i, gpu in enumerate(system_data.get("gpu", [])):
            gpu_usage_lbl = ft.Text("0 %", size=13, weight=ft.FontWeight.BOLD)
            gpu_temp_lbl = ft.Text("-- \u00b0C", size=13, weight=ft.FontWeight.BOLD)
            gpu_hist = gpu_histories[i] if i < len(gpu_histories) else deque(maxlen=60)
            if i >= len(gpu_histories):
                gpu_histories.append(gpu_hist)
            gpu_chart_cont, gpu_chart_fill, gpu_chart_line, gpu_chart_canvas = create_usage_chart(gpu_hist, "gpu")
            
            dyn_labels.append({"type": "gpu_usage", "idx": i, "label": gpu_usage_lbl, "fill_path": gpu_chart_fill, "line_path": gpu_chart_line, "canvas": gpu_chart_canvas, "history": gpu_hist})
            dyn_labels.append({"type": "gpu_temp", "idx": i, "label": gpu_temp_lbl})

            gpu_cols.extend([
                make_row("name", gpu.get("name", "N/A")),
                make_row("vendor", gpu.get("vendor", "N/A")),
                make_row("vram", f"{gpu.get('vram_gb', 0)} GB"),
                make_row("driver_version", gpu.get("driver_version", "N/A")),
                ft.Row([ft.Text(t("gpu_usage"), width=180, size=13, color=ft.Colors.GREY), gpu_usage_lbl]),
                ft.Row([ft.Text(t("temperature"), width=180, size=13, color=ft.Colors.GREY), gpu_temp_lbl]),
                gpu_chart_cont,
                ft.Divider() if i < len(system_data["gpu"])-1 else ft.Container()
            ])

        gpu_card = ft.Container(
            content=ft.Column([
                ft.Row([ft.Icon(ft.Icons.MONITOR, color=acc, size=20), ft.Text(t('gpu'), size=16, weight="bold", color=acc)]),
                ft.Divider(),
            ] + gpu_cols),
            bgcolor=bg, border_radius=10, padding=15, expand=1
        )

        # --- RAM ---
        ram = system_data.get("ram", {})
        ram_usage_lbl = ft.Text("0 %", size=13, weight=ft.FontWeight.BOLD)
        ram_avail_lbl = ft.Text(f"? / {ram.get('total_gb', 0)} GB", size=13, weight=ft.FontWeight.BOLD)
        ram_chart_cont, ram_fill_path, ram_line_path, ram_chart_canvas = create_usage_chart(ram_history, "ram")
        dyn_labels.append({"type": "ram", "usage": ram_usage_lbl, "avail": ram_avail_lbl, "total": ram.get('total_gb', 0), "fill_path": ram_fill_path, "line_path": ram_line_path, "canvas": ram_chart_canvas, "history": ram_history})

        ram_card = ft.Container(
            content=ft.Column([
                ft.Row([ft.Icon(ft.Icons.MEMORY, color=acc, size=20), ft.Text(t('ram'), size=16, weight="bold", color=acc)]),
                ft.Divider(),
                ft.Row([ft.Text(t("ram_usage"), width=180, size=13, color=ft.Colors.GREY), ram_usage_lbl]),
                ft.Row([ft.Text(t("total_available"), width=180, size=13, color=ft.Colors.GREY), ram_avail_lbl]),
                make_row("slots", f"{ram.get('used_slots', 0)} / {ram.get('total_slots', 0)}"),
                make_row("channel_mode", ram.get("channel_mode", "N/A")),
                ram_chart_cont,
            ]),
            bgcolor=bg, border_radius=10, padding=15, expand=1
        )

        # --- Disk ---
        disk_cols = []
        for i, d in enumerate(system_data.get("disks", [])):
            disk_temp = ft.Text("-- \u00b0C", size=13, weight=ft.FontWeight.BOLD)
            disk_act = ft.Text("0 %", size=13, weight=ft.FontWeight.BOLD)
            dyn_labels.append({"type": "disk_temp", "idx": i, "label": disk_temp})
            dyn_labels.append({"type": "disk_act", "idx": i, "label": disk_act})

            disk_cols.extend([
                make_row("model", d.get("model", "N/A")),
                make_row("size", f"{d.get('size_gb', 'N/A')} GB"),
                make_row("type", d.get("type", "N/A")),
                ft.Row([ft.Text(t("temperature"), width=180, size=13, color=ft.Colors.GREY), disk_temp]),
                ft.Row([ft.Text(t("disk_activity"), width=180, size=13, color=ft.Colors.GREY), disk_act]),
                ft.Container(height=10) if i < len(system_data["disks"])-1 else ft.Container()
            ])

        disk_card = ft.Container(
            content=ft.Column([
                ft.Row([ft.Icon(ft.Icons.STORAGE, color=acc, size=20), ft.Text(t('disk'), size=16, weight="bold", color=acc)]),
                ft.Divider(),
            ] + disk_cols),
            bgcolor=bg, border_radius=10, padding=15, expand=1
        )

        # --- DIAGNOSTICS LOGIC ---
        health = analyze_health(system_data, smart_data, battery_data)
        flags = detect_red_flags(system_data, smart_data)

        verdict_colors = {
            "green": ft.Colors.GREEN,
            "yellow": ft.Colors.ORANGE,
            "red": ft.Colors.RED
        }
        v_color = verdict_colors.get(health.get("overall", "green"), ft.Colors.GREEN)
        grade = health.get("grade", "A")
        
        diag_content = [
            ft.Row([
                ft.Text(t("overall_verdict"), size=18, weight="bold"),
                ft.Text(grade, size=24, weight="bold", color=v_color),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)
        ]
        
        if flags:
            diag_content.append(ft.Text(t("red_flags"), size=14, color=ALERT, weight="bold"))
            for obj in flags:
                diag_content.append(ft.Row([
                    ft.Icon(ft.Icons.WARNING_ROUNDED, color=ALERT, size=18),
                    ft.Text(f"{obj['component']}: {t(obj['message_key']).format(**obj.get('message_args', {}))}", size=13, color=ft.Colors.RED_200)
                ]))
        else:
            diag_content.append(ft.Text(t("no_red_flags"), color=ft.Colors.GREEN))

        diag_card = ft.Container(
            content=ft.Column(diag_content),
            bgcolor=bg, border_radius=10, padding=15
        )

        # SMART Table
        smart_tiles = []
        for d in smart_data:
            rows = []
            for attr in d.get("all_attrs", []):
                name = attr.get("name", "")
                desc = smart_desc(name)
                if desc:
                    name_cell = ft.DataCell(ft.Row([
                        ft.Text(name),
                        ft.Icon(ft.Icons.HELP_OUTLINE, size=14, tooltip=desc),
                    ], spacing=4, tight=True))
                else:
                    name_cell = ft.DataCell(ft.Text(name))
                rows.append(ft.DataRow(cells=[
                    ft.DataCell(ft.Text(str(attr["id"]))),
                    name_cell,
                    ft.DataCell(ft.Text(str(attr["value"]))),
                    ft.DataCell(ft.Text(str(attr["worst"]))),
                    ft.DataCell(ft.Text(str(attr["thresh"]))),
                    ft.DataCell(ft.Text(str(attr["raw"]))),
                ]))
            
            dt = ft.DataTable(
                columns=[
                    ft.DataColumn(ft.Row([ft.Text(t("smart_id")), ft.Icon(ft.Icons.HELP_OUTLINE, size=14, tooltip=smart_col_help("id"))], spacing=4, tight=True)),
                    ft.DataColumn(ft.Row([ft.Text(t("smart_attr")), ft.Icon(ft.Icons.HELP_OUTLINE, size=14, tooltip=smart_col_help("name"))], spacing=4, tight=True)),
                    ft.DataColumn(ft.Row([ft.Text(t("smart_value")), ft.Icon(ft.Icons.HELP_OUTLINE, size=14, tooltip=smart_col_help("value"))], spacing=4, tight=True)),
                    ft.DataColumn(ft.Row([ft.Text(t("smart_worst")), ft.Icon(ft.Icons.HELP_OUTLINE, size=14, tooltip=smart_col_help("worst"))], spacing=4, tight=True)),
                    ft.DataColumn(ft.Row([ft.Text(t("smart_thresh")), ft.Icon(ft.Icons.HELP_OUTLINE, size=14, tooltip=smart_col_help("thresh"))], spacing=4, tight=True)),
                    ft.DataColumn(ft.Row([ft.Text(t("smart_raw")), ft.Icon(ft.Icons.HELP_OUTLINE, size=14, tooltip=smart_col_help("raw"))], spacing=4, tight=True)),
                ],
                rows=rows,
                heading_row_height=30,
                data_row_min_height=30,
                data_row_max_height=30,
                column_spacing=15
            )
            
            dev_name = d.get('device', '')
            if dev_name.startswith('/dev/sd'):
                idx = ord(dev_name[-1]) - ord('a')
                dev_name = f"Disk {idx}"
                
            smart_tiles.append(
                ft.ExpansionTile(
                    title=ft.Text(f"S.M.A.R.T. {dev_name} ({d.get('model', '')})"), 
                    controls=[ft.Row([dt], scroll="auto")]
                )
            )
        
        smart_card = ft.Container(
            content=ft.Column(smart_tiles),
            bgcolor=bg, border_radius=10, padding=15
        )

        # Battery
        bat_card = ft.Container()
        if battery_data and battery_data.get("detected"):
            bat_card = ft.Container(
                content=ft.Column([
                    ft.Row([ft.Icon(ft.Icons.BATTERY_CHARGING_FULL, color=acc, size=20), ft.Text(t('battery'), size=16, weight="bold", color=acc)]),
                    ft.Divider(),
                    make_row("battery_status", battery_data.get("status", "Unknown")),
                    make_row("battery_charge", f"{battery_data.get('charge_pct', 0)} %"),
                    make_row("battery_wear", f"{battery_data.get('wear_pct', 0)} %"),
                    make_row("battery_cycles", battery_data.get("cycle_count", "N/A")),
                ]),
                bgcolor=bg, border_radius=10, padding=15
            )

        # Tools
        active_test = None

        def run_stress_test(e):
            nonlocal active_test
            if active_test:
                log_msg("A test is already running.")
                return
            log_msg(f"Starting {t('cpu_stress')} (60s)...")
            active_test = CpuStressTest(60, 
                lambda sec, temps, clock, usage: log_msg(f"CPU Burn: {sec}s | Max Temp: {max(temps) if temps else 0}\xb0C | Usage: {usage}%"),
                lambda res: log_msg(f"CPU Burn Finished. Passed: {res.get('passed', False)}")
            )
            active_test.start()
            def watch_test():
                nonlocal active_test
                time.sleep(65)
                active_test = None
            threading.Thread(target=watch_test).start()

        def run_disk_test(e):
            log_msg(f"Starting {t('disk_speed')}...")
            def on_done(res):
                if res.get('error'):
                    log_msg(f"Disk Test Error: {res['error']}")
                else:
                    log_msg(f"Disk Test Done. Read: {res['read_mbps']} MB/s, Write: {res['write_mbps']} MB/s")
            run_disk_speed_test("C:\\\\", size_mb=256, on_done=on_done)

        def run_ram_test(e):
            log_msg(f"Starting {t('ram_check')} (1024MB)...")
            def on_done(res):
                if res.get('error'):
                    log_msg(f"RAM Test Error: {res['error']}")
                else:
                    log_msg(f"RAM Test Done. Speed: {res['speed_mbps']} MB/s | Errors: {res['errors']}")
            run_ram_check(1024, on_done=on_done)

        def do_export(e):
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                initialfile="HardwareDiag_Report.pdf",
            )
            root.destroy()
            if path:
                try:
                    v_health = analyze_health(system_data, smart_data, battery_data)
                    v_flags = detect_red_flags(system_data, smart_data)
                    res = generate_pdf(path, system_data, verdict=v_health, red_flags=v_flags, smart_details=smart_data, battery=battery_data, lang=lang, t_fn=t)
                    if res:
                        log_msg(f"Report exported to: {path}")
                    else:
                        log_msg("Export failed or fpdf not installed")
                except Exception as ex:
                    import traceback
                    log_msg(f"Export failed: {ex}")
                    log_msg(traceback.format_exc())
            page.update()
        
        tool_card = ft.Container(
            content=ft.Column([
                ft.Row([ft.Icon(ft.Icons.BUILD, color=acc, size=20), ft.Text(t('tools'), size=16, weight="bold", color=acc)]),
                ft.Divider(),
                ft.Row([
                    ft.ElevatedButton(t("cpu_stress"), on_click=run_stress_test),
                    ft.ElevatedButton(t("disk_speed"), on_click=run_disk_test),
                    ft.ElevatedButton(t("ram_check"), on_click=run_ram_test),
                ], wrap=True),
                ft.Row([
                    ft.ElevatedButton(t("export_pdf"), on_click=do_export),
                    ft.ElevatedButton(t("copy_report"), on_click=copy_report),
                ], wrap=True)
            ]),
            bgcolor=bg, border_radius=10, padding=15
        )

        _assemble(cpu_card, gpu_card, ram_card, disk_card, diag_card, smart_card, bat_card, tool_card)

        # Apply current dynamic data immediately so lang/theme change doesn't show reset values
        try:
            dyn = collect_dynamic()
            for obj in dyn_labels:
                lbl_type = obj.get("type", "")
                try:
                    if lbl_type == "cpu_clock":
                        obj["label"].value = f"{dyn.get('current_clock_mhz', 0)} MHz"
                    elif lbl_type == "cpu_usage":
                        val = dyn.get("cpu_usage_total", 0)
                        obj["label"].value = f"{val} %"
                        update_usage_chart(obj["fill_path"], obj["line_path"], obj["canvas"], obj["history"], val)
                    elif lbl_type == "cpu_temp":
                        temps = dyn.get("cpu_temps", [])
                        val = temps[0]["value"] if temps else "--"
                        obj["label"].value = f"{val} \u00b0C"
                    elif lbl_type == "gpu_usage":
                        ld = dyn.get("gpu_load", [])
                        idx = int(obj["idx"])
                        val = ld[idx] if idx < len(ld) else 0
                        obj["label"].value = f"{val} %" if idx < len(ld) else "-- %"
                        update_usage_chart(obj["fill_path"], obj["line_path"], obj["canvas"], obj["history"], val)
                    elif lbl_type == "gpu_temp":
                        tp = dyn.get("gpu_temps", [])
                        idx = int(obj["idx"])
                        obj["label"].value = f"{tp[idx]['value']} \u00b0C" if idx < len(tp) else "-- \u00b0C"
                    elif lbl_type == "ram":
                        avail = dyn.get("ram_available_gb", 0)
                        tot = float(obj["total"])
                        obj["avail"].value = f"{avail} / {tot} GB"
                        if tot > 0:
                            val_pc = float((tot - avail) / tot * 100)
                            obj["usage"].value = f"{val_pc:.1f} %"
                            update_usage_chart(obj["fill_path"], obj["line_path"], obj["canvas"], obj["history"], val_pc)
                    elif lbl_type == "disk_act":
                        da = dyn.get("disk_activity", [])
                        idx = int(obj["idx"])
                        obj["label"].value = f"{da[idx]} %" if idx < len(da) else "-- %"
                    elif lbl_type == "disk_temp":
                        dt = dyn.get("disk_temps", [])
                        idx = int(obj["idx"])
                        obj["label"].value = f"{dt[idx]['value']} \u00b0C" if idx < len(dt) else "-- \u00b0C"
                except Exception:
                    pass
            page.update()
        except Exception:
            pass

    def _assemble(cpu_card, gpu_card, ram_card, disk_card, diag_card, smart_card, bat_card, tool_card):
        dashboard_col.controls.clear()
        dashboard_col.controls.append(ft.Row([cpu_card, gpu_card], alignment=ft.MainAxisAlignment.START, vertical_alignment=ft.CrossAxisAlignment.START))
        dashboard_col.controls.append(ft.Row([ram_card, disk_card], alignment=ft.MainAxisAlignment.START, vertical_alignment=ft.CrossAxisAlignment.START))
        dashboard_col.controls.append(diag_card)
        if bat_card.content:
            dashboard_col.controls.append(bat_card)
        if smart_card.content and hasattr(smart_card.content, "controls") and smart_card.content.controls:
            dashboard_col.controls.append(smart_card)
        dashboard_col.controls.append(tool_card)
        dashboard_col.controls.append(ft.Text("Console Output (Tests / Scans)", weight="bold"))
        dashboard_col.controls.append(console_output)
        page.update()

    def generate_dashboard():
        dashboard_col.controls.clear()
        loading_row = ft.Row([
            ft.ProgressRing(), 
            ft.Text(t("scanning") if "scanning" in T.get(lang, T["en"]) else "Scanning hardware, this may take a moment...", size=16)
        ], alignment=ft.MainAxisAlignment.CENTER)
        
        dashboard_col.controls.append(loading_row)
        page.update()

        def scan_worker():
            try:
                import pythoncom
                pythoncom.CoInitialize()
                nonlocal system_data, smart_data, battery_data
                system_data = collect_all()
                smart_data = get_deep_smart()
                battery_data = get_battery_info()
            except Exception as e:
                log_msg(f"Scan failed: {e}")
                import traceback
                log_msg(traceback.format_exc())
                return
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            try:
                loop = page.session.connection.loop
                loop.call_soon_threadsafe(build_ui)
            except Exception:
                build_ui()

        threading.Thread(target=scan_worker).start()

    generate_dashboard()

    def update_loop():
        while True:
            time.sleep(0.5)
            try:
                dyn = collect_dynamic()
                for obj in dyn_labels:
                    lbl_type = obj.get("type", "")
                    try:
                        if lbl_type == "cpu_clock":
                            obj["label"].value = f"{dyn.get('current_clock_mhz', 0)} MHz"
                            for o in [obj["label"]]:
                                o.update()
                        elif lbl_type == "cpu_usage":
                            val = dyn.get("cpu_usage_total", 0)
                            obj["label"].value = f"{val} %"
                            update_usage_chart(obj["fill_path"], obj["line_path"], obj["canvas"], obj["history"], val)
                        elif lbl_type == "cpu_temp":
                            temps = dyn.get("cpu_temps", [])
                            val = temps[0]["value"] if temps else "--"
                            obj["label"].value = f"{val} \u00b0C"
                            for o in [obj["label"]]:
                                o.update()
                        elif lbl_type == "gpu_usage":
                            ld = dyn.get("gpu_load", [])
                            idx = int(obj["idx"])
                            val = ld[idx] if idx < len(ld) else 0
                            obj["label"].value = f"{val} %" if idx < len(ld) else "-- %"
                            for o in [obj["label"]]:
                                o.update()
                            update_usage_chart(obj["fill_path"], obj["line_path"], obj["canvas"], obj["history"], val)
                        elif lbl_type == "gpu_temp":
                            tp = dyn.get("gpu_temps", [])
                            idx = int(obj["idx"])
                            obj["label"].value = f"{tp[idx]['value']} \u00b0C" if idx < len(tp) else "-- \u00b0C"
                            for o in [obj["label"]]:
                                o.update()
                        elif lbl_type == "ram":
                            avail = dyn.get("ram_available_gb", 0)
                            tot = float(obj["total"])
                            obj["avail"].value = f"{avail} / {tot} GB"
                            if tot > 0:
                                val_pc = float((tot - avail) / tot * 100)
                                obj["usage"].value = f"{val_pc:.1f} %"
                                update_usage_chart(obj["fill_path"], obj["line_path"], obj["canvas"], obj["history"], val_pc)
                            for o in [obj["avail"], obj["usage"]]:
                                o.update()
                        elif lbl_type == "disk_act":
                            da = dyn.get("disk_activity", [])
                            idx = int(obj["idx"])
                            obj["label"].value = f"{da[idx]} %" if idx < len(da) else "-- %"
                            for o in [obj["label"]]:
                                o.update()
                        elif lbl_type == "disk_temp":
                            dt = dyn.get("disk_temps", [])
                            idx = int(obj["idx"])
                            obj["label"].value = f"{dt[idx]['value']} \u00b0C" if idx < len(dt) else "-- \u00b0C"
                            for o in [obj["label"]]:
                                o.update()
                    except Exception:
                        pass
            except Exception:
                pass

    threading.Thread(target=update_loop, daemon=True).start()
    page.add(header_row, ft.Divider(), dashboard_col)

if __name__ == "__main__":
    ft.run(main)
