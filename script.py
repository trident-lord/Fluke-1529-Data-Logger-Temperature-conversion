import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import serial
import serial.tools.list_ports
import pandas as pd
import time
from datetime import datetime
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from collections import deque
import threading
import queue
import os
import math
import numpy as np
from dateutil import parser

# --- ITS-90 Conversion (PRT - Platinum Resistance Thermometer) ---
def its90_temperature(R, Rtpw=100.0):
    """
    Converts PRT Resistance (Ohms) to Temperature (°C) using ITS-90 approximation.
    """
    if Rtpw <= 0 or R is None or R <= 0:
        return float('nan')
    W = R / Rtpw
    A = 3.9083e-3
    B = -5.775e-7
    
    discriminant = A**2 - 4 * B * (1 - W)
    if discriminant < 0:
        return float('nan')
    
    return (-A + math.sqrt(discriminant)) / (2 * B)

# --- Type S Thermocouple Conversion - NIST Standard Polynomial ---
def emf_to_temperature_nist(emf_mV: float) -> float | str:
    """
    NIST Standard Method: Converts EMF (mV) to Temperature (°C) for Type S thermocouple.
    """
    if emf_mV is None or math.isnan(emf_mV):
        return float('nan')
    
    ranges = [
        (-0.235, 1.874, [
            0.0000000000E+00, 1.8494946000E+02, -8.0050406200E+01, 1.0223743000E+02,
            -1.5224859200E+02, 1.8882134300E+02, -1.5908594100E+02, 8.2302788000E+01,
            -2.3418194400E+01, 2.7978626000E+00
        ]),
        (1.874, 11.950, [
            1.2915071770E+01, 1.4662988630E+02, -1.5347134020E+01, 3.1459459730E+00,
            -4.1632578390E-01, 3.1879637710E-02, -1.2916375000E-03, 2.1834750870E-05,
            -1.4473795110E-07, 8.2112721250E-09
        ]),
        (10.332, 17.536, [
            -8.0878011170E+01, 1.6215731040E+02, -8.5368694530E+00, 4.7196869760E-01,
            -1.4416936660E-02, 2.0816188900E-04
        ]),
        (17.536, 18.693, [
            5.3338751260E+04, -1.2358922980E+04, 1.0926576130E+03, -4.2656936860E+01,
            6.2472054200E-01
        ])
    ]

    for (v_min, v_max, coeffs) in ranges:
        if (v_min - 1e-6) <= emf_mV <= (v_max + 1e-6):
            temperature = 0
            for i, c_val in enumerate(coeffs):
                temperature += c_val * (emf_mV ** i)
            return temperature
            
    if emf_mV < ranges[0][0]:
        coeffs = ranges[0][2]
        temperature = 0
        for i, c_val in enumerate(coeffs):
            temperature += c_val * (emf_mV ** i)
        return temperature
    elif emf_mV > ranges[-1][1]:
        coeffs = ranges[-1][2]
        temperature = 0
        for i, c_val in enumerate(coeffs):
            temperature += c_val * (emf_mV ** i)
        return temperature

    return float('nan')

# --- Type S Thermocouple Conversion - Custom Chart Interpolation ---
def convert_emf_to_temp_table_interpolation(measured_emf_mv: float) -> float | str:
    """
    Converts EMF (mV) to Temperature (°C) using linear interpolation from Type S calibration table.
    """
    calibration_data = [
        (0.00000, 0), (0.05514, 10), (0.11266, 20), (0.17244, 30), (0.23436, 40),
        (0.29829, 50), (0.36414, 60), (0.43179, 70), (0.50115, 80), (0.57214, 90),
        (0.64466, 100), (0.71863, 110), (0.79397, 120), (0.87062, 130), (0.94850, 140),
        (1.02755, 150), (1.10771, 160), (1.18891, 170), (1.27112, 180), (1.35427, 190),
        (1.43831, 200), (1.52321, 210), (1.60892, 220), (1.69540, 230), (1.78261, 240),
        (1.87051, 250), (1.95908, 260), (2.04828, 270), (2.13809, 280), (2.22847, 290),
        (2.31941, 300), (2.41087, 310), (2.50283, 320), (2.59528, 330), (2.68820, 340),
        (2.78157, 350), (2.87537, 360), (2.96958, 370), (3.06420, 380), (3.15921, 390),
        (3.25460, 400), (3.35035, 410), (3.44647, 420), (3.54293, 430), (3.63974, 440),
        (3.73688, 450), (3.83436, 460), (3.93215, 470), (4.03027, 480), (4.12871, 490),
        (4.22745, 500), (4.32651, 510), (4.42587, 520), (4.52554, 530), (4.62552, 540),
        (4.72581, 550), (4.82639, 560), (4.92729, 570), (5.02849, 580), (5.12999, 590),
        (5.23181, 600), (5.33394, 610), (5.43637, 620), (5.53913, 630), (5.64219, 640),
        (5.74558, 650), (5.84929, 660), (5.95332, 670), (6.05767, 680), (6.16236, 690),
        (6.26737, 700), (6.37272, 710), (6.47840, 720), (6.58442, 730), (6.69078, 740),
        (6.79748, 750), (6.90453, 760), (7.01192, 770), (7.11965, 780), (7.22773, 790),
        (7.33616, 800), (7.44493, 810), (7.55406, 820), (7.66353, 830), (7.77335, 840),
        (7.88352, 850), (7.99403, 860), (8.10489, 870), (8.21609, 880), (8.32763, 890),
        (8.43951, 900), (8.55173, 910), (8.66429, 920), (8.77718, 930), (8.89039, 940),
        (9.00394, 950), (9.11781, 960), (9.23201, 970), (9.34652, 980), (9.46136, 990),
        (9.57651, 1000), (9.69197, 1010), (9.80776, 1020), (9.92385, 1030), (10.04027, 1040),
        (10.15700, 1050), (10.27406, 1060), (10.39143, 1070), (10.50907, 1080), (10.62698, 1090),
        (10.74513, 1100), (10.86353, 1110), (10.98215, 1120), (11.10100, 1130), (11.22006, 1140),
        (11.33932, 1150), (11.45877, 1160), (11.57841, 1170), (11.69822, 1180), (11.81820, 1190),
        (11.938, 1200)
    ]

    min_emf = calibration_data[0][0]
    max_emf = calibration_data[-1][0]

    if measured_emf_mv < min_emf - 1e-6:
        emf1, temp1 = calibration_data[0]
        emf2, temp2 = calibration_data[1]
        if (emf2 - emf1) == 0: return temp1 
        return temp1 + (measured_emf_mv - emf1) * (temp2 - temp1) / (emf2 - emf1)
    
    if measured_emf_mv > max_emf + 1e-6:
        emf1, temp1 = calibration_data[-2]
        emf2, temp2 = calibration_data[-1]
        if (emf2 - emf1) == 0: return temp1 
        return temp1 + (measured_emf_mv - emf1) * (temp2 - temp1) / (emf2 - emf1)

    for i in range(len(calibration_data) - 1):
        emf1, temp1 = calibration_data[i]
        emf2, temp2 = calibration_data[i + 1]
        if emf1 <= measured_emf_mv <= emf2:
            if (emf2 - emf1) == 0: 
                return temp1
            return temp1 + (measured_emf_mv - emf1) * (temp2 - temp1) / (emf2 - emf1)
            
    return float('nan')

# --- Configuration ---
channel_configs = {
    1: {'type': 'RES', 'unit': 'O', 'enabled': None},
    2: {'type': 'RES', 'unit': 'O', 'enabled': None},
    3: {'type': 'TC', 'unit': 'MV', 'enabled': None},
    4: {'type': 'TC', 'unit': 'MV', 'enabled': None},
}

SAVE_DIR = os.path.expanduser("~/Desktop")
PLOT_MAX_POINTS = 300
PLOT_UPDATE_INTERVAL_MS = 500
SAVE_INTERVAL_RECORDS = 60
SAVE_INTERVAL_SECONDS = 300
TIMESTAMP_TIMEOUT = 2  # Timeout in seconds for grouping channel data by timestamp

new_records_buffer = []
current_record = {}  # {timestamp: {channel: {data}}}
plot_timestamps = deque(maxlen=PLOT_MAX_POINTS)
plot_data = {
    i: {
        'emf': deque(maxlen=PLOT_MAX_POINTS), 
        'temp_nist': deque(maxlen=PLOT_MAX_POINTS),
        'temp_chart': deque(maxlen=PLOT_MAX_POINTS),
        'resistance': deque(maxlen=PLOT_MAX_POINTS),
        'temp_prt': deque(maxlen=PLOT_MAX_POINTS)
    }
    for i in range(1, 5)
}

latest_values = {
    i: {
        'raw': 'N/A',
        'temp': 'N/A'
    }
    for i in range(1, 5)
}

active_plot_channel = 1
plot_type = 'temp'
separate_windows = {i: False for i in range(1, 5)}
stop_event = threading.Event()
data_queue = queue.Queue()
command_queue = queue.Queue()
ser = None
ani = None
last_save_time = 0

# --- GUI Setup ---
root = tk.Tk()
root.title("Fluke 1529 Data Logger")
root.geometry("1366x768")
root.configure(bg="#f0f0f0")
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

# Initialize BooleanVar for channel_configs after root creation
for i in range(1, 5):
    channel_configs[i]['enabled'] = tk.BooleanVar(value=True)

style = ttk.Style()
style.theme_use('clam')
style.configure("TLabel", background="#f0f0f0", font=("Arial", 11))
style.configure("TButton", padding=6, font=("Arial", 10))
style.configure("TCombobox", padding=5, font=("Arial", 10))
style.configure("Card.TFrame", background="#ffffff", relief="solid", borderwidth=1)

main_frame = ttk.Frame(root, padding=10)
main_frame.pack(fill="both", expand=True)
left_panel = ttk.Frame(main_frame, style="Card.TFrame")
left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
right_panel = ttk.Frame(main_frame, style="Card.TFrame")
right_panel.grid(row=0, column=1, sticky="nsew")
main_frame.columnconfigure(1, weight=3)

real_time_frame = ttk.LabelFrame(left_panel, text="Real-Time Values (Chart-based for TC)", padding=10)
real_time_frame.pack(fill="x", pady=5, padx=5)

value_labels = {}
for i in range(1, 5):
    ttk.Label(real_time_frame, text=f"Channel {i}:", font=("Arial", 11, "bold")).grid(row=i-1, column=0, sticky="w", padx=5)
    raw_label = ttk.Label(real_time_frame, text="N/A", font=("Arial", 12))
    raw_label.grid(row=i-1, column=1, sticky="w", padx=5)
    temp_label = ttk.Label(real_time_frame, text="N/A", font=("Arial", 14), foreground="#e74c3c")
    temp_label.grid(row=i-1, column=2, sticky="w", padx=10)
    value_labels[i] = {'raw': raw_label, 'temp': temp_label}

controls_frame = ttk.LabelFrame(left_panel, text="Controls", padding=10)
controls_frame.pack(fill="x", pady=5, padx=5)

conn_frame = ttk.Frame(controls_frame)
conn_frame.pack(fill="x")
com_ports = [port.device for port in serial.tools.list_ports.comports()]
com_port_var = tk.StringVar(value=com_ports[0] if com_ports else "")
ttk.Label(conn_frame, text="COM Port:").grid(row=0, column=0, sticky="w", padx=5)
com_port_combo = ttk.Combobox(conn_frame, textvariable=com_port_var, values=com_ports, state="readonly")
com_port_combo.grid(row=0, column=1, padx=5)

ttk.Label(conn_frame, text="Baud:").grid(row=0, column=2, sticky="w", padx=5)
baud_rate_var = tk.IntVar(value=9600)
baud_rate_combo = ttk.Combobox(conn_frame, textvariable=baud_rate_var, values=[9600, 19200, 38400, 57600, 115200], state="readonly")
baud_rate_combo.grid(row=0, column=3, padx=5)

btn_frame = ttk.Frame(controls_frame)
btn_frame.pack(fill="x", pady=5)
start_button = ttk.Button(btn_frame, text="Start Logging", command=lambda: start_logging())
start_button.pack(side="left", padx=2)
stop_button = ttk.Button(btn_frame, text="Stop Logging", command=lambda: stop_logging(), state="disabled")
stop_button.pack(side="left", padx=2)
calibrate_button = ttk.Button(btn_frame, text="Calibrate Time", command=lambda: calibrate_time())
calibrate_button.pack(side="left", padx=2)

settings_frame = ttk.LabelFrame(left_panel, text="Settings", padding=10)
settings_frame.pack(fill="x", pady=5, padx=5)

ttk.Label(settings_frame, text="Measure Period:").grid(row=0, column=0, sticky="w", pady=2)
meas_period_var = tk.StringVar(value="1s")
meas_period_combo = ttk.Combobox(settings_frame, textvariable=meas_period_var, values=["0.1s", "0.2s", "0.5s", "1s", "2s", "5s", "10s", "30s", "1min", "2min", "5min", "10min", "30min", "1hr"], state="readonly")
meas_period_combo.grid(row=0, column=1, sticky="w", pady=2)

ttk.Label(settings_frame, text="Save Directory:").grid(row=1, column=0, sticky="w", pady=2)
save_dir_var = tk.StringVar(value=SAVE_DIR)
ttk.Entry(settings_frame, textvariable=save_dir_var, state="readonly").grid(row=1, column=1, sticky="w", pady=2)
ttk.Button(settings_frame, text="Browse", command=lambda: browse_directory(save_dir_var)).grid(row=1, column=2, padx=5, pady=2)

channel_enable_frame = ttk.LabelFrame(left_panel, text="Channel Enable", padding=10)
channel_enable_frame.pack(fill="x", pady=5, padx=5)
for i in range(1, 5):
    ttk.Checkbutton(channel_enable_frame, text=f"Enable Channel {i}", variable=channel_configs[i]['enabled']).grid(row=i-1, column=0, sticky="w", padx=5, pady=2)

unit_frame = ttk.LabelFrame(left_panel, text="Unit Settings", padding=10)
unit_frame.pack(fill="x", pady=5, padx=5)

unit_vars = {i: tk.StringVar(value=channel_configs[i]['unit']) for i in range(1, 5)}
for i in range(1, 5):
    ttk.Label(unit_frame, text=f"Ch {i} Unit:").grid(row=i-1, column=0, sticky="e", padx=5, pady=2)
    unit_combo = ttk.Combobox(unit_frame, textvariable=unit_vars[i], values=["O", "MV"], state="readonly")
    unit_combo.grid(row=i-1, column=1, padx=5, pady=2)
    unit_combo.bind('<<ComboboxSelected>>', lambda event, ch=i: send_unit_command(ch))

plot_frame = ttk.Frame(right_panel, padding=10)
plot_frame.pack(fill="both", expand=True)

toggle_frame = ttk.Frame(plot_frame)
toggle_frame.pack(fill="x", pady=5)
ttk.Label(toggle_frame, text="Select Channel for Main Plot:").pack(side="left", padx=5)
channel_buttons = {}
for i in range(1, 5):
    btn = ttk.Button(toggle_frame, text=f"Ch {i}", command=lambda ch=i: set_active_channel(ch))
    btn.pack(side="left", padx=2)
    channel_buttons[i] = btn

sub_toggle_frame = ttk.Frame(plot_frame)
sub_toggle_frame.pack(fill="x", pady=5)
ttk.Button(sub_toggle_frame, text="Raw vs Time", command=lambda: set_plot_type("raw")).pack(side="left", padx=2)
ttk.Button(sub_toggle_frame, text="Temp vs Time", command=lambda: set_plot_type("temp")).pack(side="left", padx=2)
ttk.Button(sub_toggle_frame, text="All Temp vs Time", command=lambda: show_all_channels()).pack(side="left", padx=2)

checkbox_frame = ttk.Frame(plot_frame)
checkbox_frame.pack(fill="x", pady=5)
check_vars = {i: tk.BooleanVar() for i in range(1, 5)}
for i in range(1, 5):
    ttk.Checkbutton(checkbox_frame, text=f"Open Ch {i} in Separate Window", variable=check_vars[i],
                    command=lambda ch=i: toggle_separate_window(ch)).pack(side="left", padx=2)

fig, ax = plt.subplots(figsize=(10, 6))
fig.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.15)
ax.grid(True)
ax.set_xlabel("Time")
ax.set_ylabel("Value")
ax.tick_params(axis='x', rotation=45)

lines = {}
for i in range(1, 5):
    if channel_configs[i]['type'] == 'RES':
        lines[f'ch{i}_prt'] = ax.plot([], [], label=f'Ch {i} PRT Temp (°C)', color=f'C{i-1}')[0]
    elif channel_configs[i]['type'] == 'TC':
        lines[f'ch{i}_nist'] = ax.plot([], [], label=f'Ch {i} TC Temp (NIST) (°C)', color=f'C{i-1}', linestyle='-')[0]
        lines[f'ch{i}_chart'] = ax.plot([], [], label=f'Ch {i} TC Temp (Chart) (°C)', color=f'C{i-1}', linestyle='--')[0]
    
for key in lines:
    lines[key].set_visible(False)

canvas = FigureCanvasTkAgg(fig, master=plot_frame)
canvas.get_tk_widget().pack(fill="both", expand=True)
toolbar = NavigationToolbar2Tk(canvas, plot_frame)
toolbar.update()
toolbar.pack(side="bottom", fill="x")

window_figures = {i: None for i in range(1, 5)}
window_canvases = {i: None for i in range(1, 5)}
window_axes = {i: None for i in range(1, 5)}
window_lines = {i: {} for i in range(1, 5)}

status_var = tk.StringVar(value="Ready. Select COM port and press Start.")
ttk.Frame(root, padding=5).pack(fill="x", side="bottom")
ttk.Label(root, textvariable=status_var).pack(side="left", padx=10)

# --- Core Logic ---
def serial_reader_thread():
    """Reads data from the serial port and puts it into a queue for processing."""
    global ser
    try:
        ser = serial.Serial(com_port_var.get(), baud_rate_var.get(), timeout=1)
        status_var.set(f"Connected to {com_port_var.get()}")
        send_scpi_command(f"MEAS:PER {meas_period_var.get().replace('s','').replace('min','m').replace('hr','h')}")
        for ch in range(1, 5):
            if channel_configs[ch]['enabled'].get():
                send_unit_command(ch)
    except serial.SerialException as e:
        status_var.set(f"Connection failed: {e}")
        stop_event.set()
        return

    while not stop_event.is_set():
        try:
            while not command_queue.empty():
                cmd = command_queue.get()
                ser.write((cmd + '\n').encode())
                time.sleep(0.1)
            if ser.in_waiting:
                line = ser.readline().decode(errors='ignore').strip()
                print(f"Raw serial data: {line}")  # Debug: Print raw serial data
                line_parts = line.split()
                if len(line_parts) >= 5:
                    channel = int(line_parts[0])
                    if not channel_configs[channel]['enabled'].get():
                        continue
                    raw_val_str = line_parts[1]
                    if raw_val_str == '........':
                        continue
                    try:
                        raw_val = float(raw_val_str)
                    except ValueError:
                        print(f"Error parsing serial data: could not convert string to float: '{raw_val_str}' - Line: {line}")
                        status_var.set(f"Data parsing error for Channel {channel}")
                        continue
                    unit = line_parts[2]
                    timestamp_str = f"{line_parts[4]} {line_parts[3]}"
                    data_queue.put({'channel': channel, 'raw_val': raw_val, 'unit': unit, 'timestamp': timestamp_str})
                    status_var.set(f"Received data for Channel {channel}: {raw_val} {unit}")
                    print(f"Queued data: Channel {channel}, Value {raw_val}, Unit {unit}, Timestamp {timestamp_str}")
                else:
                    print(f"Invalid serial data format: {line}")
        except (ValueError, IndexError) as e:
            print(f"Error parsing serial data: {e} - Line: {line}")
            status_var.set(f"Data parsing error: {e}")
            continue
        except serial.SerialException:
            status_var.set("Serial error. Attempting to reconnect...")
            time.sleep(2)
        except Exception as e:
            status_var.set(f"Unexpected serial thread error: {e}")
            break

    if ser and ser.is_open:
        ser.close()
    status_var.set("Disconnected")

def animate(frame):
    """Updates plots in real-time by processing data from the queue."""
    global last_save_time, current_record

    # Process data from the queue
    while not data_queue.empty():
        data = data_queue.get()
        channel = data['channel']
        timestamp_str = data['timestamp']
        try:
            timestamp = datetime.strptime(timestamp_str, '%d/%m/%Y %H:%M:%S')
        except ValueError:
            try:
                timestamp = datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S')
            except ValueError:
                try:
                    timestamp = parser.parse(timestamp_str)
                except Exception as e:
                    print(f"Failed to parse timestamp: {timestamp_str} — {e}")
                    continue
        unit = data['unit']
        raw_val = data['raw_val']

        # Initialize current_record for this timestamp if not present
        timestamp_key = timestamp.strftime("%Y-%m-%d %H:%M:%S")
        if timestamp_key not in current_record:
            current_record[timestamp_key] = {'channels': {}, 'receive_time': time.time()}

        # Update plot_data and latest_values immediately for real-time display
        emf, temp_nist, temp_chart, resistance, temp_prt = float('nan'), float('nan'), float('nan'), float('nan'), float('nan')
        if channel_configs[channel]['type'] == 'RES':
            resistance = raw_val
            try:
                temp_prt = its90_temperature(resistance)
            except ValueError as e:
                temp_prt = float('nan')
                print(f"PRT conversion error for Ch{channel}: {e}")
            plot_data[channel]['resistance'].append(resistance)
            plot_data[channel]['temp_prt'].append(temp_prt)
            latest_values[channel]['raw'] = f"{resistance:.4f} Ω"
            latest_values[channel]['temp'] = f"{temp_prt:.4f} °C" if not math.isnan(temp_prt) else "N/A"
            current_record[timestamp_key]['channels'][channel] = {'resistance': resistance, 'temp_prt': temp_prt}
        elif channel_configs[channel]['type'] == 'TC':
            emf = raw_val
            try:
                temp_nist = emf_to_temperature_nist(emf)
            except ValueError as e:
                temp_nist = float('nan')
                print(f"TC NIST conversion error for Ch{channel}: {e}")
            try:
                temp_chart = convert_emf_to_temp_table_interpolation(emf)
            except ValueError as e:
                temp_chart = float('nan')
                print(f"TC Chart conversion error for Ch{channel}: {e}")
            difference = temp_chart - temp_nist if not math.isnan(temp_nist) and not math.isnan(temp_chart) else float('nan')
            plot_data[channel]['emf'].append(emf)
            plot_data[channel]['temp_nist'].append(temp_nist)
            plot_data[channel]['temp_chart'].append(temp_chart)
            latest_values[channel]['raw'] = f"{emf:.4f} mV"
            latest_values[channel]['temp'] = f"{temp_chart:.4f} °C" if not math.isnan(temp_chart) else "N/A"
            current_record[timestamp_key]['channels'][channel] = {'emf': emf, 'temp_nist': temp_nist, 'temp_chart': temp_chart, 'difference': difference}

        if not plot_timestamps or timestamp > plot_timestamps[-1]:
            plot_timestamps.append(timestamp)

    # Check for complete or timed-out records
    enabled_channels = [ch for ch in range(1, 5) if channel_configs[ch]['enabled'].get()]
    current_time = time.time()
    for timestamp_key in list(current_record.keys()):
        record_data = current_record[timestamp_key]
        received_channels = set(record_data['channels'].keys())
        if all(ch in received_channels for ch in enabled_channels) or \
           (current_time - record_data['receive_time'] >= TIMESTAMP_TIMEOUT):
            # Create a single record for this timestamp
            record = [timestamp_key]
            for ch in range(1, 5):
                if ch in record_data['channels']:
                    if channel_configs[ch]['type'] == 'RES':
                        record.extend([
                            record_data['channels'][ch]['resistance'],
                            record_data['channels'][ch]['temp_prt']
                        ])
                    elif channel_configs[ch]['type'] == 'TC':
                        record.extend([
                            record_data['channels'][ch]['emf'],
                            record_data['channels'][ch]['temp_nist'],
                            record_data['channels'][ch]['temp_chart'],
                            record_data['channels'][ch]['difference']
                        ])
                else:
                    if channel_configs[ch]['type'] == 'RES':
                        record.extend([float('nan'), float('nan')])
                    elif channel_configs[ch]['type'] == 'TC':
                        record.extend([float('nan'), float('nan'), float('nan'), float('nan')])
            new_records_buffer.append(record)
            print(f"Processed record for timestamp {timestamp_key}: {record}")
            del current_record[timestamp_key]

    update_real_time_labels()
    update_main_plot()
    if new_records_buffer and (len(new_records_buffer) >= SAVE_INTERVAL_RECORDS or (current_time - last_save_time >= SAVE_INTERVAL_SECONDS)):
        save_to_excel(new_records_buffer)
        new_records_buffer.clear()
        last_save_time = current_time

    for ch in range(1, 5):
        if separate_windows[ch] and window_figures[ch]:
            update_separate_window(ch)

def update_main_plot():
    """Updates the main matplotlib plot based on active_plot_channel and plot_type."""
    for key in lines:
        lines[key].set_visible(False)

    x_data = list(plot_timestamps)
    
    if not x_data:
        ax.set_xlim(datetime.now() - pd.Timedelta(minutes=1), datetime.now())
        ax.relim()
        ax.autoscale_view(scaley=True)
        fig.canvas.draw_idle()
        return

    if plot_type == 'temp':
        ax.set_ylabel("Temperature (°C)")
    else:
        unit_str = "Ω" if channel_configs[active_plot_channel]['type'] == 'RES' else "mV"
        ax.set_ylabel(f"Raw Value ({unit_str})")
    
    if plot_type == 'temp':
        if channel_configs[active_plot_channel]['type'] == 'RES':
            y_data = list(plot_data[active_plot_channel]['temp_prt'])
            if len(y_data) > 0 and len(x_data) >= len(y_data):
                x_data_subset = x_data[-len(y_data):]
                lines[f'ch{active_plot_channel}_prt'].set_data(x_data_subset, y_data)
                lines[f'ch{active_plot_channel}_prt'].set_visible(True)
        elif channel_configs[active_plot_channel]['type'] == 'TC':
            y_data_nist = list(plot_data[active_plot_channel]['temp_nist'])
            y_data_chart = list(plot_data[active_plot_channel]['temp_chart'])
            if len(y_data_nist) > 0 and len(x_data) >= len(y_data_nist):
                x_data_subset = x_data[-len(y_data_nist):]
                lines[f'ch{active_plot_channel}_nist'].set_data(x_data_subset, y_data_nist)
                lines[f'ch{active_plot_channel}_nist'].set_visible(True)
            if len(y_data_chart) > 0 and len(x_data) >= len(y_data_chart):
                x_data_subset = x_data[-len(y_data_chart):]
                lines[f'ch{active_plot_channel}_chart'].set_data(x_data_subset, y_data_chart)
                lines[f'ch{active_plot_channel}_chart'].set_visible(True)
    else:
        if channel_configs[active_plot_channel]['type'] == 'RES':
            y_data = list(plot_data[active_plot_channel]['resistance'])
            if len(y_data) > 0 and len(x_data) >= len(y_data):
                x_data_subset = x_data[-len(y_data):]
                lines[f'ch{active_plot_channel}_prt'].set_data(x_data_subset, y_data)
                lines[f'ch{active_plot_channel}_prt'].set_visible(True)
        elif channel_configs[active_plot_channel]['type'] == 'TC':
            y_data = list(plot_data[active_plot_channel]['emf'])
            if len(y_data) > 0 and len(x_data) >= len(y_data):
                x_data_subset = x_data[-len(y_data):]
                lines[f'ch{active_plot_channel}_nist'].set_data(x_data_subset, y_data)
                lines[f'ch{active_plot_channel}_nist'].set_visible(True)
    
    if len(x_data) > 1:
        ax.set_xlim(x_data[0], x_data[-1])
    else:
        ax.set_xlim(x_data[0] - pd.Timedelta(seconds=1), x_data[0] + pd.Timedelta(seconds=1))
    ax.legend(handles=[lines[key] for key in lines if lines[key].get_visible()])
    ax.relim()
    ax.autoscale_view(scaley=True)
    fig.canvas.draw_idle()

def update_real_time_labels():
    """Updates the Tkinter labels displaying the latest sensor values."""
    for ch, labels in value_labels.items():
        if channel_configs[ch]['enabled'].get():
            labels['raw'].config(text=latest_values[ch]['raw'])
            labels['temp'].config(text=latest_values[ch]['temp'])
        else:
            labels['raw'].config(text="Disabled")
            labels['temp'].config(text="Disabled")

def save_to_excel(records):
    """Saves collected data records to an Excel file."""
    if not records:
        return
    date_str = datetime.now().strftime("%Y%m%d")
    excel_file = os.path.join(save_dir_var.get(), f"fluke_1529_{date_str}.xlsx")
    
    # Dynamically generate columns based on current channel configurations
    columns = ['Timestamp']
    for i in range(1, 5):
        if channel_configs[i]['enabled'].get():
            if channel_configs[i]['type'] == 'RES':
                columns.append(f'Ch{i} PRT Resistance (Ω)')
                columns.append(f'Ch{i} PRT Temperature (°C)')
            elif channel_configs[i]['type'] == 'TC':
                columns.append(f'Ch{i} TC EMF (mV)')
                columns.append(f'Ch{i} TC Temp (NIST) (°C)')
                columns.append(f'Ch{i} TC Temp (Chart) (°C)')
                columns.append(f'Ch{i} Difference (Chart - NIST) (°C)')
        else:
            if channel_configs[i]['type'] == 'RES':
                columns.append(f'Ch{i} PRT Resistance (Ω)')
                columns.append(f'Ch{i} PRT Temperature (°C)')
            elif channel_configs[i]['type'] == 'TC':
                columns.append(f'Ch{i} TC EMF (mV)')
                columns.append(f'Ch{i} TC Temp (NIST) (°C)')
                columns.append(f'Ch{i} TC Temp (Chart) (°C)')
                columns.append(f'Ch{i} Difference (Chart - NIST) (°C)')

    try:
        new_df = pd.DataFrame(records, columns=columns)
        
        if os.path.exists(excel_file):
            existing_df = pd.read_excel(excel_file)
            updated_df = pd.concat([existing_df, new_df], ignore_index=True)
        else:
            updated_df = new_df
        
        updated_df.to_excel(excel_file, index=False, engine='openpyxl')
        status_var.set(f"Saved {len(records)} records to {os.path.basename(excel_file)}")
    except Exception as e:
        status_var.set(f"Save failed: {e}")
        messagebox.showerror("Error", f"Excel save error: {e}")

def start_logging():
    """Initializes and starts data logging."""
    global new_records_buffer, plot_timestamps, plot_data, last_save_time, stop_event, data_queue, ani, current_record
    
    COM_PORT = com_port_var.get()
    if not COM_PORT:
        messagebox.showerror("Input Error", "Please select a valid COM port.")
        return
    try:
        BAUD_RATE = int(baud_rate_var.get())
        if BAUD_RATE <= 0: raise ValueError("Baud Rate must be positive")
    except ValueError as e:
        messagebox.showerror("Input Error", f"Invalid Baud Rate: {e}")
        return

    try:
        temp_ser = serial.Serial(COM_PORT, BAUD_RATE, timeout=1)
        temp_ser.close()
    except serial.SerialException as se:
        messagebox.showerror("Serial Error", f"COM port {COM_PORT} is not available or in use: {se}")
        return

    new_records_buffer.clear()
    current_record.clear()
    plot_timestamps.clear()
    plot_data = {
        i: {
            'emf': deque(maxlen=PLOT_MAX_POINTS), 
            'temp_nist': deque(maxlen=PLOT_MAX_POINTS),
            'temp_chart': deque(maxlen=PLOT_MAX_POINTS),
            'resistance': deque(maxlen=PLOT_MAX_POINTS),
            'temp_prt': deque(maxlen=PLOT_MAX_POINTS)
        }
        for i in range(1, 5)
    }
    for ch in range(1, 5):
        latest_values[ch] = {'raw': 'N/A', 'temp': 'N/A'}
    
    last_save_time = time.time()
    stop_event.clear()
    data_queue = queue.Queue()

    status_var.set("Starting serial connection...")
    serial_thread = threading.Thread(target=serial_reader_thread, daemon=True)
    serial_thread.start()

    ani = FuncAnimation(fig, animate, interval=PLOT_UPDATE_INTERVAL_MS, cache_frame_data=False)
    canvas.draw()

    start_button.config(state="disabled")
    calibrate_button.config(state="normal")
    stop_button.config(state="normal")
    for ch in range(1, 5):
        channel_buttons[ch].config(state="disabled" if not channel_configs[ch]['enabled'].get() else "normal")
    status_var.set("Logging started")
    messagebox.showinfo("Info", "Logging started. Data will be saved to Excel periodically.")

def stop_logging():
    """Stops data logging and cleans up resources."""
    global ani, ser
    stop_event.set()
    if ani:
        ani.event_source.stop()
    
    start_button.config(state="normal")
    stop_button.config(state="disabled")
    calibrate_button.config(state="normal")
    for ch in range(1, 5):
        channel_buttons[ch].config(state="normal")
    
    if new_records_buffer:
        save_to_excel(new_records_buffer)
        new_records_buffer.clear()
    
    # Process any remaining partial records
    for timestamp_key in list(current_record.keys()):
        record_data = current_record[timestamp_key]
        record = [timestamp_key]
        for ch in range(1, 5):
            if ch in record_data['channels']:
                if channel_configs[ch]['type'] == 'RES':
                    record.extend([
                        record_data['channels'][ch]['resistance'],
                        record_data['channels'][ch]['temp_prt']
                    ])
                elif channel_configs[ch]['type'] == 'TC':
                    record.extend([
                        record_data['channels'][ch]['emf'],
                        record_data['channels'][ch]['temp_nist'],
                        record_data['channels'][ch]['temp_chart'],
                        record_data['channels'][ch]['difference']
                    ])
            else:
                if channel_configs[ch]['type'] == 'RES':
                    record.extend([float('nan'), float('nan')])
                elif channel_configs[ch]['type'] == 'TC':
                    record.extend([float('nan'), float('nan'), float('nan'), float('nan')])
        new_records_buffer.append(record)
        print(f"Processed final record for timestamp {timestamp_key}: {record}")
    if new_records_buffer:
        save_to_excel(new_records_buffer)
        new_records_buffer.clear()
    
    if ser and ser.is_open:
        ser.close()
    status_var.set("Logging stopped")
    
    for ch in range(1, 5):
        if window_figures[ch]:
            window_figures[ch].get_tk_widget().master.destroy()
            window_figures[ch] = None
            window_canvases[ch] = None
            window_axes[ch] = None
            window_lines[ch] = {}
            separate_windows[ch] = False

def calibrate_time():
    """Sends SCPI commands to synchronize instrument time with PC time."""
    if ser and ser.is_open:
        now = datetime.now()
        command_queue.put(f"SYST:DATE {now.year},{now.month:02d},{now.day:02d}")
        command_queue.put(f"SYST:TIME {now.hour:02d},{now.minute:02d},{now.second:02d}")
        messagebox.showinfo(
            "Calibration",
            f"Sent: SYST:DATE {now.year}/{now.month:02d}/{now.day:02d}, "
            f"SYST:TIME {now.hour:02d}:{now.minute:02d}:{now.second:02d}"
        )
        status_var.set("Time calibration commands sent")
    else:
        messagebox.showerror("Error", "Not connected to instrument. Start logging first.")


def send_scpi_command(command):
    """Sends a general SCPI command to the instrument."""
    if ser and ser.is_open:
        command_queue.put(command)
        status_var.set(f"Sent: {command}")
    else:
        status_var.set("Not connected. Command will apply on next start.")

def send_unit_command(channel):
    """Sends SCPI command to set the unit for a specific channel (1-4)."""
    if not isinstance(channel, int) or channel < 1 or channel > 4:
        status_var.set(f"Error: Invalid channel {channel}. Must be 1-4.")
        return
    unit = unit_vars[channel].get()
    if ser and ser.is_open:
        command_queue.put(f"UNIT:CHAN{channel} {unit}")
        channel_configs[channel]['unit'] = unit
        channel_configs[channel]['type'] = 'RES' if unit == 'O' else 'TC'
        status_var.set(f"Sent: Set Channel {channel} unit to {unit}. Internal type updated.")
    else:
        status_var.set("Not connected. Unit change will apply on next start.")
        channel_configs[channel]['unit'] = unit
        channel_configs[channel]['type'] = 'RES' if unit == 'O' else 'TC'

def set_active_channel(channel):
    """Sets the active channel for the main plot and redraws."""
    global active_plot_channel
    if channel_configs[channel]['enabled'].get():
        active_plot_channel = channel
        update_main_plot()

def set_plot_type(ptype):
    """Sets the plot type ('raw' or 'temp') for the main plot and redraws."""
    global plot_type
    plot_type = ptype
    update_main_plot()

def show_all_channels():
    """Displays all enabled channels' temperature data on the main plot."""
    global plot_type
    plot_type = 'temp'
    
    x_data = list(plot_timestamps)
    if not x_data:
        ax.set_xlim(datetime.now() - pd.Timedelta(minutes=1), datetime.now())
        ax.relim()
        ax.autoscale_view(scaley=True)
        fig.canvas.draw_idle()
        return

    ax.set_ylabel("Temperature (°C)")

    for key in lines:
        lines[key].set_visible(False)

    for ch in range(1, 5):
        if channel_configs[ch]['enabled'].get():
            if channel_configs[ch]['type'] == 'RES':
                y_data = list(plot_data[ch]['temp_prt'])
                if len(y_data) > 0 and len(x_data) >= len(y_data):
                    x_data_subset = x_data[-len(y_data):]
                    lines[f'ch{ch}_prt'].set_data(x_data_subset, y_data)
                    lines[f'ch{ch}_prt'].set_visible(True)
            elif channel_configs[ch]['type'] == 'TC':
                y_data_nist = list(plot_data[ch]['temp_nist'])
                y_data_chart = list(plot_data[ch]['temp_chart'])
                if len(y_data_nist) > 0 and len(x_data) >= len(y_data_nist):
                    x_data_subset = x_data[-len(y_data_nist):]
                    lines[f'ch{ch}_nist'].set_data(x_data_subset, y_data_nist)
                    lines[f'ch{ch}_nist'].set_visible(True)
                if len(y_data_chart) > 0 and len(x_data) >= len(y_data_chart):
                    x_data_subset = x_data[-len(y_data_chart):]
                    lines[f'ch{ch}_chart'].set_data(x_data_subset, y_data_chart)
                    lines[f'ch{ch}_chart'].set_visible(True)
    
    if len(x_data) > 1:
        ax.set_xlim(x_data[0], x_data[-1])
    else:
        ax.set_xlim(x_data[0] - pd.Timedelta(seconds=1), x_data[0] + pd.Timedelta(seconds=1))
    ax.legend(handles=[lines[key] for key in lines if lines[key].get_visible()])
    ax.relim()
    ax.autoscale_view(scaley=True)
    fig.canvas.draw_idle()

def toggle_separate_window(channel):
    """Opens or closes a separate plot window for a specific channel."""
    global window_figures, window_canvases, separate_windows, window_axes, window_lines
    
    if check_vars[channel].get() and channel_configs[channel]['enabled'].get():
        if not window_figures[channel]:
            window = tk.Toplevel(root)
            window.title(f"Channel {channel} Plot")
            window.geometry("800x600")
            window.protocol("WM_DELETE_WINDOW", lambda ch=channel: close_separate_window_callback(ch))
            
            fig_ch = plt.Figure(figsize=(8, 6), dpi=100)
            ax_ch = fig_ch.add_subplot(111)
            ax_ch.grid(True)
            ax_ch.set_xlabel("Time")
            ax_ch.set_ylabel("Value")
            ax_ch.tick_params(axis='x', rotation=45)
            
            if channel_configs[channel]['type'] == 'RES':
                window_lines[channel]['prt'] = ax_ch.plot([], [], label=f'Ch {channel} PRT Temp (°C)', color='C0')[0]
            elif channel_configs[channel]['type'] == 'TC':
                window_lines[channel]['nist'] = ax_ch.plot([], [], label=f'Ch {channel} TC Temp (NIST) (°C)', color='C0', linestyle='-')[0]
                window_lines[channel]['chart'] = ax_ch.plot([], [], label=f'Ch {channel} TC Temp (Chart) (°C)', color='C0', linestyle='--')[0]

            ax_ch.legend()

            canvas_ch = FigureCanvasTkAgg(fig_ch, master=window)
            canvas_ch.get_tk_widget().pack(fill="both", expand=True)
            toolbar_ch = NavigationToolbar2Tk(canvas_ch, window)
            toolbar_ch.update()
            toolbar_ch.pack(side="bottom", fill="x")
            
            window_figures[channel] = fig_ch
            window_canvases[channel] = canvas_ch
            window_axes[channel] = ax_ch
            separate_windows[channel] = True
            
            update_separate_window(channel)
    else:
        close_separate_window_callback(channel)

def update_separate_window(channel):
    """Updates the content of a specific separate plot window."""
    if window_figures[channel] and separate_windows[channel]:
        ax_ch = window_axes[channel]
        x_data = list(plot_timestamps)
        
        if not x_data:
            ax_ch.set_xlim(datetime.now() - pd.Timedelta(minutes=1), datetime.now())
            ax_ch.relim()
            ax_ch.autoscale_view(scaley=True)
            window_canvases[channel].draw_idle()
            return

        if plot_type == 'temp':
            ax_ch.set_ylabel("Temperature (°C)")
        else:
            unit_str = "Ω" if channel_configs[channel]['type'] == 'RES' else "mV"
            ax_ch.set_ylabel(f"Raw Value ({unit_str})")

        for key in window_lines[channel]:
            window_lines[channel][key].set_visible(False)

        if plot_type == 'temp':
            if channel_configs[channel]['type'] == 'RES':
                y_data = list(plot_data[channel]['temp_prt'])
                if len(y_data) > 0 and len(x_data) >= len(y_data):
                    x_data_subset = x_data[-len(y_data):]
                    window_lines[channel]['prt'].set_data(x_data_subset, y_data)
                    window_lines[channel]['prt'].set_visible(True)
            elif channel_configs[channel]['type'] == 'TC':
                y_data_nist = list(plot_data[channel]['temp_nist'])
                y_data_chart = list(plot_data[channel]['temp_chart'])
                if len(y_data_nist) > 0 and len(x_data) >= len(y_data_nist):
                    x_data_subset = x_data[-len(y_data_nist):]
                    window_lines[channel]['nist'].set_data(x_data_subset, y_data_nist)
                    window_lines[channel]['nist'].set_visible(True)
                if len(y_data_chart) > 0 and len(x_data) >= len(y_data_chart):
                    x_data_subset = x_data[-len(y_data_chart):]
                    window_lines[channel]['chart'].set_data(x_data_subset, y_data_chart)
                    window_lines[channel]['chart'].set_visible(True)
        else:
            if channel_configs[channel]['type'] == 'RES':
                y_data = list(plot_data[channel]['resistance'])
                if len(y_data) > 0 and len(x_data) >= len(y_data):
                    x_data_subset = x_data[-len(y_data):]
                    window_lines[channel]['prt'].set_data(x_data_subset, y_data)
                    window_lines[channel]['prt'].set_visible(True)
            elif channel_configs[channel]['type'] == 'TC':
                y_data = list(plot_data[channel]['emf'])
                if len(y_data) > 0 and len(x_data) >= len(y_data):
                    x_data_subset = x_data[-len(y_data):]
                    window_lines[channel]['nist'].set_data(x_data_subset, y_data)
                    window_lines[channel]['nist'].set_visible(True)

        if len(x_data) > 1:
            ax_ch.set_xlim(x_data[0], x_data[-1])
        else:
            ax_ch.set_xlim(x_data[0] - pd.Timedelta(seconds=1), x_data[0] + pd.Timedelta(seconds=1))
        ax_ch.legend(handles=[window_lines[channel][key] for key in window_lines[channel] if window_lines[channel][key].get_visible()])
        ax_ch.relim()
        ax_ch.autoscale_view(scaley=True)
        window_canvases[channel].draw_idle()

def close_separate_window_callback(channel):
    """Callback function for when a separate window is closed by the user."""
    if window_figures[channel]:
        window_figures[channel].get_tk_widget().master.destroy()
        window_figures[channel] = None
        window_canvases[channel] = None
        window_axes[channel] = None
        window_lines[channel] = {}
        separate_windows[channel] = False
        check_vars[channel].set(False)

def browse_directory(var):
    """Opens a file dialog to select a save directory."""
    new_dir = filedialog.askdirectory(initialdir=var.get(), title="Select Save Directory")
    if new_dir:
        var.set(new_dir)

def on_closing():
    """Handles the application closing event."""
    if messagebox.askokcancel("Quit", "Quit application?"):
        stop_logging()
        root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()
