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

# --- ITS-90 Conversion ---
def its90_temperature(R, Rtpw=100.0):
    if Rtpw <= 0 or R is None or R <= 0:
        return float('nan')
    W = R / Rtpw
    A = 3.9083e-3
    B = -5.775e-7
    discriminant = A**2 - 4 * B * (1 - W)
    if discriminant < 0:
        return float('nan')
    return (-A + math.sqrt(discriminant)) / (2 * B)

def emf_to_temperature(emf_mV):
    if emf_mV is None:
        return float('nan')
    
    #subranges and corresponding coefficients
    ranges = [
        (-0.235, 1.874, [0.0000000E+00, 1.84949460E+02, -8.00504062E+01, 1.02237430E+02,
                         -1.52248592E+02, 1.88821343E+02, -1.59085941E+02, 8.23027880E+01,
                         -2.34181944E+01, 2.79786260E+00]),
        
        (1.874, 11.950, [1.291507177E+01, 1.466298863E+02, -1.534713402E+01, 3.145945973E+00,
                         -4.163257839E-01, 3.187963771E-02, -1.291637500E-03, 2.183475087E-05,
                         -1.447379511E-07, 8.211272125E-09]),
        
        (10.332, 17.536, [-8.087801117E+01, 1.621573104E+02, -8.536869453E+00, 4.719686976E-01,
                          -1.441693666E-02, 2.081618890E-04, 0.0, 0.0, 0.0, 0.0]),
        
        (17.536, 18.693, [5.333875126E+04, -1.235892298E+04, 1.092657613E+03, -4.265693686E+01,
                          6.247205420E-01, 0.0, 0.0, 0.0, 0.0, 0.0])
    ]

    for (v_min, v_max, coeffs) in ranges:
        if v_min <= emf_mV <= v_max:
            temperature = 0
            for i, d in enumerate(coeffs):
                temperature += d * (emf_mV ** i)
            return temperature+0.02
    raise ValueError("EMF out of range for Type S thermocouple.")


# --- Configuration ---
columns = ['Timestamp', 'Thermocouple EMF', 'Thermocouple Temperature', 'PRT Resistance', 'PRT Temperature']
SAVE_DIR = os.path.expanduser("~/Desktop")
PLOT_MAX_POINTS = 300
PLOT_UPDATE_INTERVAL_MS = 500
SAVE_INTERVAL_RECORDS = 60
SAVE_INTERVAL_SECONDS = 300
new_records_buffer = []

plot_timestamps = deque(maxlen=PLOT_MAX_POINTS)
plot_data = {i: {'emf': deque(maxlen=PLOT_MAX_POINTS), 'temp_prt': deque(maxlen=PLOT_MAX_POINTS),
                 'resistance': deque(maxlen=PLOT_MAX_POINTS), 'temp_tc': deque(maxlen=PLOT_MAX_POINTS)}
             for i in range(1, 5)}
channel_configs = {i: {'type': 'RES', 'unit': 'O', 'enabled': True} for i in range(1, 5)}
latest_values = {i: {'emf': 'N/A', 'temp_prt': 'N/A', 'resistance': 'N/A', 'temp_tc': 'N/A'} for i in range(1, 5)}
active_plot_channel = 1
plot_type = 'temp'  # 'raw' for EMF/resistance, 'temp' for temperature
separate_windows = {i: False for i in range(1, 5)}  # Track separate window states

# Threading
stop_event = threading.Event()
data_queue = queue.Queue()
command_queue = queue.Queue()
ser = None
ani = None
last_save_time = 0  # Initialize globally

# --- GUI Setup ---
root = tk.Tk()
root.title("Fluke 1529 Data Logger")
root.geometry("1366x768")
root.configure(bg="#f0f0f0")
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

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


# Real-Time Values
real_time_frame = ttk.LabelFrame(left_panel, text="Real-Time Values", padding=10)
real_time_frame.pack(fill="x", pady=5, padx=5)

value_labels = {i: {'temp': ttk.Label(real_time_frame, text="N/A", font=("Arial", 14), foreground="#e74c3c")}
                for i in range(1, 5) if channel_configs[i]['enabled']}
for i, labels in value_labels.items():
    ttk.Label(real_time_frame, text=f"Channel {i} Temp:", font=("Arial", 11, "bold")).grid(row=i-1, column=0, sticky="w", padx=5)
    labels['temp'].grid(row=i-1, column=1, sticky="w", padx=10)

# Controls
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

# Settings
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

# Unit Setting
unit_frame = ttk.LabelFrame(left_panel, text="Unit Settings", padding=10)
unit_frame.pack(fill="x", pady=5, padx=5)
unit_vars = {i: tk.StringVar(value="O") for i in range(1, 5)}
for i in range(1, 5):
    ttk.Label(unit_frame, text=f"Ch {i} Unit:").grid(row=i-1, column=0, sticky="e", padx=5, pady=2)
    unit_combo = ttk.Combobox(unit_frame, textvariable=unit_vars[i], values=["O", "MV"], state="readonly")
    unit_combo.grid(row=i-1, column=1, padx=5, pady=2)
    unit_combo.bind('<<ComboboxSelected>>', lambda event, ch=i: send_unit_command(ch))

# Plot Toggles and Checkboxes
plot_frame = ttk.Frame(right_panel, padding=10)
plot_frame.pack(fill="both", expand=True)

toggle_frame = ttk.Frame(plot_frame)
toggle_frame.pack(fill="x", pady=5)
ttk.Label(toggle_frame, text="Select Channel:").pack(side="left", padx=5)
channel_buttons = {}
for i in range(1, 5):
    if channel_configs[i]['enabled']:
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
    if channel_configs[i]['enabled']:
        ttk.Checkbutton(checkbox_frame, text=f"Open Ch {i} in Separate Window", variable=check_vars[i],
                       command=lambda ch=i: toggle_separate_window(ch)).pack(side="left", padx=2)

# Main Plot
fig, ax = plt.subplots()
fig.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.1)
lines = {i: ax.plot([], [], label=f'Ch {i}')[0] for i in range(1, 5)}
ax.grid(True)
ax.set_xlabel("Time")
ax.set_ylabel("Value")
ax.legend()

canvas = FigureCanvasTkAgg(fig, master=plot_frame)
canvas.get_tk_widget().pack(fill="both", expand=True)
toolbar = NavigationToolbar2Tk(canvas, plot_frame)
toolbar.pack(side="bottom", fill="x")

# Separate Window Figures
window_figures = {i: None for i in range(1, 5)}
window_canvases = {i: None for i in range(1, 5)}

status_var = tk.StringVar(value="Ready. Select COM port and press Start.")
ttk.Frame(root, padding=5).pack(fill="x", side="bottom")
ttk.Label(root, textvariable=status_var).pack(side="left", padx=10)

# --- Core Logic ---
def serial_reader_thread():
    global ser
    try:
        ser = serial.Serial(com_port_var.get(), baud_rate_var.get(), timeout=1)
        status_var.set(f"Connected to {com_port_var.get()}")
        send_scpi_command(f"MEAS:PER {meas_period_var.get().replace('s','').replace('min','m').replace('hr','h')}")
    except serial.SerialException as e:
        status_var.set(f"Connection failed: {e}")
        stop_event.set()
        return

    while not stop_event.is_set():
        try:
            while not command_queue.empty():
                ser.write((command_queue.get() + '\n').encode())
                time.sleep(0.1)
            if ser.in_waiting:
                line = ser.readline().decode(errors='ignore').strip().split()
                if len(line) >= 5:
                    channel = int(line[0])
                    raw_val = float(line[1])
                    unit = line[2]
                    timestamp = f"{line[4]} {line[3]}"
                    data_queue.put({'channel': channel, 'raw_val': raw_val, 'unit': unit, 'timestamp': timestamp})
                    status_var.set(f"Received data for Channel {channel}: {raw_val} {unit}")
        except (ValueError, IndexError):
            continue
        except serial.SerialException:
            status_var.set("Serial error. Reconnecting...")
            time.sleep(2)

    if ser.is_open:
        ser.close()
    status_var.set("Disconnected")

def animate(frame):
    global last_save_time
    if not data_queue.empty():
        data = data_queue.get()
        channel = data['channel']
        timestamp = datetime.strptime(data['timestamp'], '%d/%m/%Y %H:%M:%S')
        unit = unit_vars[channel].get()

        if not plot_timestamps or timestamp > plot_timestamps[-1]:
            plot_timestamps.append(timestamp)
        if unit == 'O':  # PRT (Resistance)
            emf = float('nan')
            resistance = data['raw_val']
            temp_prt = its90_temperature(resistance)
            temp_tc = float('nan')
        else:  # Thermocouple (EMF)
            emf = data['raw_val']
            resistance = float('nan')
            temp_prt = float('nan')
            temp_tc = emf_to_temperature(emf)

        plot_data[channel]['emf'].append(emf)
        plot_data[channel]['temp_prt'].append(temp_prt)
        plot_data[channel]['resistance'].append(resistance)
        plot_data[channel]['temp_tc'].append(temp_tc)
        latest_values[channel] = {'emf': f"{emf:.4f} mV" if not math.isnan(emf) else "N/A",
                                 'temp_prt': f"{temp_prt:.4f} °C" if not math.isnan(temp_prt) else "N/A",
                                 'resistance': f"{resistance:.4f} Ω" if not math.isnan(resistance) else "N/A",
                                 'temp_tc': f"{temp_tc:.4f} °C" if not math.isnan(temp_tc) else "N/A"}
        new_records_buffer.append([data['timestamp'], emf, temp_tc, resistance, temp_prt])

    update_real_time_labels()
    for ch in range(1, 5):
        if ch == active_plot_channel and channel_configs[ch]['enabled']:
            data_key = 'temp' if plot_type == 'temp' else 'raw'
            y_data = plot_data[ch]['temp_prt' if data_key == 'temp' else 'resistance'] if unit_vars[ch].get() == 'O' else plot_data[ch]['temp_tc' if data_key == 'temp' else 'emf']
            y_values = list(y_data)
            x_data = list(plot_timestamps)
            min_length = min(len(x_data), len(y_values))
            if min_length > 0:
                lines[ch].set_data(x_data[:min_length], [d if d is not None else float('nan') for d in y_values[:min_length]])
                lines[ch].set_visible(True)
                ax.set_ylabel("Temperature (°C)" if plot_type == 'temp' else f"Raw {unit_vars[ch].get() == 'O' and 'Ω' or 'mV'}")
            else:
                lines[ch].set_visible(False)
        else:
            lines[ch].set_visible(False)
    if len(plot_timestamps) > 1:
        ax.set_xlim(plot_timestamps[0], plot_timestamps[-1])
    else:
        ax.set_xlim(datetime.now() - pd.Timedelta(minutes=1), datetime.now())
    ax.relim()
    ax.autoscale_view(scaley=True)
    canvas.draw_idle()

    current_time = time.time()
    if len(new_records_buffer) >= SAVE_INTERVAL_RECORDS or (new_records_buffer and current_time - last_save_time >= SAVE_INTERVAL_SECONDS):
        save_to_excel(new_records_buffer)
        new_records_buffer.clear()
        last_save_time = current_time

    # Update separate windows
    for ch in range(1, 5):
        if separate_windows[ch] and window_figures[ch]:
            update_separate_window(ch)

def update_real_time_labels():
    for ch, labels in value_labels.items():
        labels['temp'].config(text=latest_values[ch]['temp_prt'] if unit_vars[ch].get() == 'O' else latest_values[ch]['temp_tc'])

def save_to_excel(records):
    if not records:
        return
    date_str = datetime.now().strftime("%Y%m%d")
    excel_file = os.path.join(save_dir_var.get(), f"fluke_1529_{date_str}.xlsx")
    temp_df = pd.DataFrame(records, columns=columns)
    try:
        if os.path.exists(excel_file):
            existing_df = pd.read_excel(excel_file)
            updated_df = pd.concat([existing_df, temp_df], ignore_index=True)
        else:
            updated_df = temp_df
        updated_df.to_excel(excel_file, index=False, engine='openpyxl')
        status_var.set(f"Saved {len(records)} records to {os.path.basename(excel_file)}")
    except Exception as e:
        status_var.set(f"Save failed: {e}")
        messagebox.showerror("Error", f"Excel save error: {e}")

def start_logging():
    global ser, new_records_buffer, plot_timestamps, plot_data, last_save_time, stop_event, data_queue, ani
    global SAVE_INTERVAL_RECORDS, SAVE_INTERVAL_SECONDS, PLOT_MAX_POINTS, PLOT_UPDATE_INTERVAL_MS

    # Validate parameters
    COM_PORT = com_port_var.get()
    if not COM_PORT or COM_PORT == "":
        messagebox.showerror("Input Error", "Please select a valid COM port.")
        return
    try:
        BAUD_RATE = int(baud_rate_var.get())
        if BAUD_RATE <= 0:
            raise ValueError("Baud Rate must be positive")
        SAVE_INTERVAL_RECORDS = int(60)  # Default, adjust if variable exists
        if SAVE_INTERVAL_RECORDS <= 0:
            raise ValueError("Save Interval Records must be positive")
        SAVE_INTERVAL_SECONDS = int(300)  # Default, adjust if variable exists
        if SAVE_INTERVAL_SECONDS <= 0:
            raise ValueError("Save Interval Seconds must be positive")
        PLOT_MAX_POINTS = int(300)  # Default, adjust if variable exists
        if PLOT_MAX_POINTS <= 0:
            raise ValueError("Plot Max Points must be positive")
        PLOT_UPDATE_INTERVAL_MS = int(500)  # Default, adjust if variable exists
        if PLOT_UPDATE_INTERVAL_MS <= 0:
            raise ValueError("Plot Update Interval must be positive")
    except ValueError as e:
        messagebox.showerror("Input Error", f"Invalid input: {e}")
        return

    # Check if COM port is connected
    try:
        ser = serial.Serial(COM_PORT, BAUD_RATE, timeout=1)
        ser.close()  # Close immediately after checking
    except serial.SerialException as se:
        messagebox.showerror("Serial Error", f"COM port {COM_PORT} is not available or in use: {se}")
        return

    # Reset data structures
    new_records_buffer.clear()
    plot_timestamps = deque(maxlen=PLOT_MAX_POINTS)
    plot_data = {i: {'emf': deque(maxlen=PLOT_MAX_POINTS), 'temp_prt': deque(maxlen=PLOT_MAX_POINTS),
                     'resistance': deque(maxlen=PLOT_MAX_POINTS), 'temp_tc': deque(maxlen=PLOT_MAX_POINTS)}
                 for i in range(1, 5)}
    latest_values = {i: {'emf': 'N/A', 'temp_prt': 'N/A', 'resistance': 'N/A', 'temp_tc': 'N/A'} for i in range(1, 5)}
    last_save_time = time.time()
    stop_event.clear()
    data_queue = queue.Queue()

    # Start serial thread
    status_var.set("Starting serial connection...")
    serial_thread = threading.Thread(target=serial_reader_thread, daemon=True)
    serial_thread.start()

    # Start animation
    ani = FuncAnimation(fig, animate, interval=PLOT_UPDATE_INTERVAL_MS, cache_frame_data=False)
    canvas.draw()

    # Update UI
    start_button.config(state="disabled")
    calibrate_button.config(state="normal")
    stop_button.config(state="normal")
    status_var.set("Logging started")
    messagebox.showinfo("Info", "Logging started.")

def stop_logging():
    stop_event.set()
    if ani:
        ani.event_source.stop()
    start_button.config(state="normal")
    stop_button.config(state="disabled")
    calibrate_button.config(state="normal")
    if new_records_buffer:
        save_to_excel(new_records_buffer)
        new_records_buffer.clear()
    if ser and ser.is_open:
        ser.close()
    status_var.set("Logging stopped")
    # Close separate windows
    for ch in range(1, 5):
        if window_figures[ch]:
            window_figures[ch].get_tk_widget().master.destroy()
            window_figures[ch] = None
            window_canvases[ch] = None
            separate_windows[ch] = False

def calibrate_time():
    if ser and ser.is_open:
        now = datetime.now()
        command_queue.put(f"SYST:DATE {now.day:02d},{now.month:02d},{now.year}")
        command_queue.put(f"SYST:TIME {now.hour:02d},{now.minute:02d},{now.second:02d}")
        messagebox.showinfo("Calibration", f"Sent: SYST:DATE {now.day:02d}/{now.month:02d}/{now.year}, SYST:TIME {now.hour:02d}:{now.minute:02d}:{now.second:02d}")
        status_var.set("Time calibration sent")
    else:
        messagebox.showerror("Error", "Not connected")

def send_unit_command(channel):
    if ser and ser.is_open:
        unit = unit_vars[channel].get()
        command_queue.put(f"UNIT:CHAN{channel} {unit}")
        channel_configs[channel]['unit'] = unit
        status_var.set(f"Set Channel {channel} unit to {unit}")

def set_active_channel(channel):
    global active_plot_channel
    active_plot_channel = channel
    animate(0)  # Force redraw

def set_plot_type(ptype):
    global plot_type
    plot_type = ptype
    animate(0)  # Force redraw

def show_all_channels():
    for ch in range(1, 5):
        if channel_configs[ch]['enabled']:
            y_data = plot_data[ch]['temp_prt'] if unit_vars[ch].get() == 'O' else plot_data[ch]['temp_tc']
            y_values = list(y_data)
            x_data = list(plot_timestamps)
            min_length = min(len(x_data), len(y_values))
            if min_length > 0:
                lines[ch].set_data(x_data[:min_length], [d if d is not None else float('nan') for d in y_values[:min_length]])
                lines[ch].set_visible(True)
            else:
                lines[ch].set_visible(False)
    ax.set_ylabel("Temperature (°C)")
    ax.legend()
    canvas.draw()

def toggle_separate_window(channel):
    global window_figures, window_canvases, separate_windows
    if check_vars[channel].get():
        if not window_figures[channel]:
            window = tk.Toplevel(root)
            window.title(f"Channel {channel} Plot")
            window.protocol("WM_DELETE_WINDOW", lambda ch=channel: close_separate_window(ch))
            fig_ch = plt.Figure(figsize=(5, 4), dpi=100)
            ax_ch = fig_ch.add_subplot(111)
            ax_ch.grid(True)
            ax_ch.set_xlabel("Time")
            ax_ch.set_ylabel("Value")
            line_ch = ax_ch.plot([], [], label=f'Ch {channel}')[0]
            canvas_ch = FigureCanvasTkAgg(fig_ch, master=window)
            canvas_ch.get_tk_widget().pack(fill="both", expand=True)
            toolbar_ch = NavigationToolbar2Tk(canvas_ch, window)
            toolbar_ch.pack(side="bottom", fill="x")
            window_figures[channel] = fig_ch
            window_canvases[channel] = canvas_ch
            separate_windows[channel] = True
            update_separate_window(channel)
    else:
        close_separate_window(channel)

def update_separate_window(channel):
    if window_figures[channel] and separate_windows[channel]:
        ax_ch = window_figures[channel].axes[0]
        y_data = plot_data[channel]['temp_prt'] if plot_type == 'temp' and unit_vars[channel].get() == 'O' else plot_data[channel]['temp_tc'] if plot_type == 'temp' else plot_data[channel]['resistance'] if unit_vars[channel].get() == 'O' else plot_data[channel]['emf']
        y_values = list(y_data)
        x_data = list(plot_timestamps)
        min_length = min(len(x_data), len(y_values))
        if min_length > 0:
            ax_ch.lines[0].set_data(x_data[:min_length], [d if d is not None else float('nan') for d in y_values[:min_length]])
            ax_ch.set_visible(True)
            ax_ch.set_ylabel("Temperature (°C)" if plot_type == 'temp' else f"Raw {unit_vars[channel].get() == 'O' and 'Ω' or 'mV'}")
        else:
            ax_ch.lines[0].set_visible(False)
        if len(plot_timestamps) > 1:
            ax_ch.set_xlim(plot_timestamps[0], plot_timestamps[-1])
        else:
            ax_ch.set_xlim(datetime.now() - pd.Timedelta(minutes=1), datetime.now())
        ax_ch.relim()
        ax_ch.autoscale_view(scaley=True)
        window_canvases[channel].draw_idle()

def close_separate_window(channel):
    if window_figures[channel]:
        window_figures[channel].get_tk_widget().master.destroy()
        window_figures[channel] = None
        window_canvases[channel] = None
        separate_windows[channel] = False
        check_vars[channel].set(False)

def browse_directory(var):
    new_dir = filedialog.askdirectory(initialdir=var.get(), title="Select Save Directory")
    if new_dir:
        var.set(new_dir)

def send_scpi_command(command):
    if ser and ser.is_open:
        command_queue.put(command)

def on_closing():
    if messagebox.askokcancel("Quit", "Quit application?"):
        stop_logging()
        root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()