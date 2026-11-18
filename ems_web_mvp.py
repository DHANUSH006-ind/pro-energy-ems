import random
import sqlite3
from datetime import datetime
import openpyxl
from nicegui import ui

# Try importing Modbus & Serial libraries for real hardware mode
try:
    import serial.tools.list_ports
    from pymodbus.client import ModbusSerialClient
    MODBUS_AVAILABLE = True
except ImportError:
    MODBUS_AVAILABLE = False

# ==========================================
# 1. SQLITE HISTORIAN DATABASE SETUP
# ==========================================
DB_FILE = "ems_historian.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS telemetry (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT,
            mode TEXT,
            slave_id INTEGER,
            voltage_v REAL,
            current_a REAL,
            power_kw REAL,
            power_factor REAL,
            energy_kwh REAL,
            water_lpm REAL,
            air_bar REAL,
            gas_m3h REAL,
            steam_bar REAL,
            status TEXT
        )
    ''')
    conn.commit()
    conn.close()

init_db()

# ==========================================
# 2. APPLICATION STATE
# ==========================================
def get_com_ports():
    if not MODBUS_AVAILABLE:
        return ["No COM Port Found"]
    ports = [p.device for p in serial.tools.list_ports.comports()]
    return ports if ports else ["No COM Port Found"]

class EMSState:
    def __init__(self):
        self.running = False
        self.mode = "Simulation"
        self.com_port = get_com_ports()[0]
        self.baudrate = 9600
        self.slave_id = 1
        self.poll_interval = 2.0
        
        # Live electrical readings
        self.voltage = 230.0
        self.current = 12.5
        self.power_kw = 2.76
        self.pf = 0.96
        self.energy_kwh = 2847.0
        
        # WAGES readings
        self.water_lpm = 125.4
        self.air_bar = 7.1
        self.gas_m3h = 42.0
        self.steam_bar = 8.5
        
        self.alert_status = "NORMAL"
        self.alert_message = "All WAGES & Electrical parameters within nominal set points."
        
        # Chart history buffers (last 20 points)
        self.timestamps = []
        self.hist_elec = []
        self.hist_water = []
        self.hist_air = []
        self.hist_steam = []

state = EMSState()

# ==========================================
# 3. DATA ACQUISITION (SIMULATION + MODBUS)
# ==========================================
def poll_data():
    if not state.running:
        return

    now_str = datetime.now().strftime("%H:%M:%S")
    full_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if state.mode == "Simulation":
        state.voltage = round(random.uniform(226.0, 236.0), 1)
        state.current = round(random.uniform(11.0, 18.5), 2)
        state.pf = round(random.uniform(0.93, 0.99), 2)
        state.power_kw = round((state.voltage * state.current * state.pf) / 1000.0, 2)
        state.energy_kwh = round(state.energy_kwh + (state.power_kw * (state.poll_interval / 3600.0)), 3)
        
        state.water_lpm = round(random.uniform(115.0, 145.0), 1)
        state.air_bar = round(random.uniform(6.4, 7.8), 2)
        state.gas_m3h = round(random.uniform(38.0, 48.0), 1)
        state.steam_bar = round(random.uniform(7.8, 9.2), 2)
    else:
        if MODBUS_AVAILABLE and "COM" in state.com_port:
            client = ModbusSerialClient(
                port=state.com_port,
                baudrate=int(state.baudrate),
                parity='N',
                stopbits=1,
                bytesize=8,
                timeout=1
            )
            if client.connect():
                try:
                    res = client.read_holding_registers(address=0, count=4, slave=int(state.slave_id))
                    if not res.isError():
                        state.voltage = round(res.registers[0] / 10.0, 1)
                        state.current = round(res.registers[1] / 100.0, 2)
                        state.pf = round(res.registers[2] / 100.0, 2) if res.registers[2] > 0 else 0.95
                        state.power_kw = round((state.voltage * state.current * state.pf) / 1000.0, 2)
                finally:
                    client.close()

    # Evaluate Set Points & Alerts
    if state.pf < 0.95:
        state.alert_status = "WARNING"
        state.alert_message = f"Power Factor Warning: {state.pf} is below target threshold (> 0.95)!"
    elif state.voltage > 245.0 or state.voltage < 210.0:
        state.alert_status = "CRITICAL"
        state.alert_message = f"Critical Voltage Deviation: {state.voltage} V outside safe operating band!"
    else:
        state.alert_status = "NORMAL"
        state.alert_message = "All WAGES & Electrical parameters within nominal set points."

    # Save to SQLite Historian
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        INSERT INTO telemetry (timestamp, mode, slave_id, voltage_v, current_a, power_kw, power_factor, energy_kwh, water_lpm, air_bar, gas_m3h, steam_bar, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (full_ts, state.mode, state.slave_id, state.voltage, state.current, state.power_kw, state.pf, state.energy_kwh, state.water_lpm, state.air_bar, state.gas_m3h, state.steam_bar, state.alert_status))
    conn.commit()
    conn.close()

    # Update rolling chart buffers
    state.timestamps.append(now_str)
    state.hist_elec.append(round(state.power_kw * 25, 1))
    state.hist_water.append(state.water_lpm)
    state.hist_air.append(round(state.air_bar * 10, 1))
    state.hist_steam.append(round(state.steam_bar * 12, 1))

    if len(state.timestamps) > 20:
        state.timestamps.pop(0)
        state.hist_elec.pop(0)
        state.hist_water.pop(0)
        state.hist_air.pop(0)
        state.hist_steam.pop(0)

    refresh_ui()

# ==========================================
# 4. EXCEL EXPORT (PURE PYTHON OPENPYXL)
# ==========================================
def export_excel():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT timestamp, mode, slave_id, voltage_v, current_a, power_kw, power_factor, energy_kwh, water_lpm, air_bar, gas_m3h, steam_bar, status FROM telemetry ORDER BY id DESC LIMIT 5000")
    rows = c.fetchall()
    conn.close()

    if not rows:
        ui.notify("No logged data available to export yet. Click START first!", type="warning")
        return

    filename = f"ProEnergy_EMS_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "WAGES & Power Telemetry"
    ws.append(["Timestamp", "Mode", "Slave ID", "Voltage (V)", "Current (A)", "Active Power (kW)", "Power Factor", "Total Energy (kWh)", "Water Flow (L/min)", "Air Pressure (bar)", "Gas Flow (m3/hr)", "Steam Pressure (bar)", "Alert Status"])
    for r in rows:
        ws.append(list(r))
    wb.save(filename)
    ui.download(filename)
    ui.notify(f"Exported {len(rows)} records to {filename}", type="positive")

# ==========================================
# 5. PROFESSIONAL LIGHT THEME & CLEAN UI
# ==========================================
ui.add_head_html('<style>body { background-color: #f8fafc; font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }</style>')

# Header with Official Pro Energy Logo and Professional Styling
with ui.header().classes("bg-white border-b border-slate-200 px-6 py-3 flex items-center justify-between shadow-sm"):
    with ui.row().classes("items-center gap-4"):
        ui.image("https://i.ibb.co/68H8Lp8w/logo.png").style("height: 42px; width: auto; object-fit: contain;")
        with ui.column().classes("gap-0"):
            ui.label("INTEGRATED EMS & SCADA PLATFORM").classes("text-xs font-bold tracking-wider text-slate-400 uppercase")
            ui.label("Pro Energy Automations | Enterprise Dashboard").classes("text-sm font-semibold text-slate-800")
    with ui.row().classes("items-center gap-3"):
        status_badge = ui.badge("SYSTEM STANDBY", color="grey").classes("text-xs px-3 py-1 font-semibold")
        ui.button("EXPORT EXCEL", icon="download", on_click=export_excel).props("color=indigo outline size=sm")

with ui.column().classes("w-full max-w-7xl mx-auto p-6 gap-5"):

    # Alert Banner
    alert_card = ui.card().classes("w-full bg-slate-50 border border-slate-200 shadow-none px-4 py-3 rounded-xl")
    with alert_card:
        with ui.row().classes("items-center gap-2"):
            alert_icon = ui.icon("check_circle", color="green", size="20px")
            alert_text = ui.label(state.alert_message).classes("text-sm font-medium text-slate-700")

    # Connection & Control Bar
    available_ports = get_com_ports()
    state.com_port = available_ports[0]

    with ui.card().classes("w-full bg-white border border-slate-200 shadow-sm p-4 rounded-xl"):
        with ui.row().classes("w-full items-end justify-between flex-wrap gap-4"):
            with ui.row().classes("items-end gap-4 flex-wrap"):
                ui.select(["Simulation", "Real Meter (RS-485 Modbus)"], value="Simulation", label="Data Source").classes("w-56").bind_value(state, "mode")
                ui.select(available_ports, value=state.com_port, label="COM Port").classes("w-40").bind_value(state, "com_port")
                ui.select([9600, 19200, 38400, 115200], value=9600, label="Baud Rate").classes("w-32").bind_value(state, "baudrate")
                ui.select({2.0: "2s (Demo)", 10.0: "10s (Standard)"}, value=2.0, label="Refresh Rate").classes("w-36").bind_value(state, "poll_interval")

            def toggle_run():
                state.running = not state.running
                if state.running:
                    start_btn.set_text("STOP ACQUISITION")
                    start_btn.props("color=red icon=stop")
                    status_badge.set_text("LIVE POLLING ACTIVE")
                    status_badge.props("color=green")
                    poll_data()
                else:
                    start_btn.set_text("START ACQUISITION")
                    start_btn.props("color=indigo icon=play_arrow")
                    status_badge.set_text("SYSTEM STANDBY")
                    status_badge.props("color=grey")

            start_btn = ui.button("START ACQUISITION", icon="play_arrow", on_click=toggle_run).props("color=indigo")

    # Electrical KPI Cards
    ui.label("ELECTRICAL SUB-METERING").classes("text-xs font-bold tracking-wider text-slate-500 uppercase mt-2 mb-1")
    with ui.grid(columns=5).classes("w-full gap-4"):
        with ui.card().classes("bg-white border border-slate-200 shadow-sm p-4 rounded-xl"):
            ui.label("PHASE VOLTAGE").classes("text-xs text-slate-500 font-medium")
            lbl_volt = ui.label("230.0 V").classes("text-2xl font-bold text-slate-900")
        with ui.card().classes("bg-white border border-slate-200 shadow-sm p-4 rounded-xl"):
            ui.label("LOAD CURRENT").classes("text-xs text-slate-500 font-medium")
            lbl_curr = ui.label("12.50 A").classes("text-2xl font-bold text-slate-900")
        with ui.card().classes("bg-white border border-slate-200 shadow-sm p-4 rounded-xl"):
            ui.label("ACTIVE POWER").classes("text-xs text-slate-500 font-medium")
            lbl_kw = ui.label("2.76 kW").classes("text-2xl font-bold text-slate-900")
        with ui.card().classes("bg-white border border-slate-200 shadow-sm p-4 rounded-xl"):
            ui.label("POWER FACTOR").classes("text-xs text-slate-500 font-medium")
            lbl_pf = ui.label("0.96").classes("text-2xl font-bold text-slate-900")
        with ui.card().classes("bg-white border border-slate-200 shadow-sm p-4 rounded-xl"):
            ui.label("TOTAL ENERGY").classes("text-xs text-slate-500 font-medium")
            lbl_kwh = ui.label("2847.0 kWh").classes("text-2xl font-bold text-slate-900")

    # WAGES Utility Cards
    ui.label("WAGES UTILITY MONITORING").classes("text-xs font-bold tracking-wider text-slate-500 uppercase mt-2 mb-1")
    with ui.grid(columns=4).classes("w-full gap-4"):
        with ui.card().classes("bg-white border border-slate-200 shadow-sm p-4 rounded-xl"):
            ui.label("WATER FLOW").classes("text-xs text-slate-500 font-medium")
            lbl_water = ui.label("125.4 L/min").classes("text-xl font-bold text-slate-900")
        with ui.card().classes("bg-white border border-slate-200 shadow-sm p-4 rounded-xl"):
            ui.label("COMPRESSED AIR").classes("text-xs text-slate-500 font-medium")
            lbl_air = ui.label("7.10 bar").classes("text-xl font-bold text-slate-900")
        with ui.card().classes("bg-white border border-slate-200 shadow-sm p-4 rounded-xl"):
            ui.label("NATURAL GAS").classes("text-xs text-slate-500 font-medium")
            lbl_gas = ui.label("42.0 m³/hr").classes("text-xl font-bold text-slate-900")
        with ui.card().classes("bg-white border border-slate-200 shadow-sm p-4 rounded-xl"):
            ui.label("PROCESS STEAM").classes("text-xs text-slate-500 font-medium")
            lbl_steam = ui.label("8.50 bar").classes("text-xl font-bold text-slate-900")

    # Live Multi-Utility Trend Chart
    with ui.card().classes("w-full bg-white border border-slate-200 shadow-sm p-4 mt-2 rounded-xl"):
        ui.label("LIVE WAGES & POWER TREND").classes("text-sm font-bold text-slate-900 mb-2")
        chart = ui.echart({
            "backgroundColor": "transparent",
            "tooltip": {"trigger": "axis"},
            "legend": {"data": ["Electricity Index", "Water (L/min)", "Air (bar x10)", "Steam (bar x12)"], "textStyle": {"color": "#334155"}},
            "xAxis": {"type": "category", "data": [], "axisLabel": {"color": "#64748b"}},
            "yAxis": {"type": "value", "axisLabel": {"color": "#64748b"}, "splitLine": {"lineStyle": {"color": "#f1f5f9"}}},
            "series": [
                {"name": "Electricity Index", "type": "line", "smooth": True, "data": [], "itemStyle": {"color": "#6366f1"}},
                {"name": "Water (L/min)", "type": "line", "smooth": True, "data": [], "itemStyle": {"color": "#0ea5e9"}},
                {"name": "Air (bar x10)", "type": "line", "smooth": True, "data": [], "itemStyle": {"color": "#10b981"}},
                {"name": "Steam (bar x12)", "type": "line", "smooth": True, "data": [], "itemStyle": {"color": "#f59e0b"}},
            ]
        }).classes("w-full h-72")

def refresh_ui():
    lbl_volt.set_text(f"{state.voltage} V")
    lbl_curr.set_text(f"{state.current:.2f} A")
    lbl_kw.set_text(f"{state.power_kw:.2f} kW")
    lbl_pf.set_text(f"{state.pf:.2f}")
    lbl_kwh.set_text(f"{state.energy_kwh:.2f} kWh")
    
    lbl_water.set_text(f"{state.water_lpm} L/min")
    lbl_air.set_text(f"{state.air_bar:.2f} bar")
    lbl_gas.set_text(f"{state.gas_m3h} m³/hr")
    lbl_steam.set_text(f"{state.steam_bar:.2f} bar")

    alert_text.set_text(state.alert_message)
    if state.alert_status == "NORMAL":
        alert_card.classes(replace="w-full bg-slate-50 border border-slate-200 shadow-none px-4 py-3 rounded-xl")
        alert_icon.props("name=check_circle color=green")
    else:
        alert_card.classes(replace="w-full bg-amber-50 border border-amber-300 shadow-none px-4 py-3 rounded-xl")
        alert_icon.props("name=warning color=amber")

    chart.options["xAxis"]["data"] = list(state.timestamps)
    chart.options["series"][0]["data"] = list(state.hist_elec)
    chart.options["series"][1]["data"] = list(state.hist_water)
    chart.options["series"][2]["data"] = list(state.hist_air)
    chart.options["series"][3]["data"] = list(state.hist_steam)
    chart.update()

# Timer loop
ui.timer(2.0, poll_data)

# Run server on port 7860 (Ready for Render cloud deployment)
ui.run(title="Pro Energy EMS & SCADA", host="0.0.0.0", port=7860, reload=False)
