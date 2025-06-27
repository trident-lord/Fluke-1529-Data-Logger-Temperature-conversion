# Fluke 1529 Chub-E Data Logger
Copyright (c) 2025 trident-lord

A sleek and robust Python application designed for real-time temperature monitoring and data logging with the Fluke 1529 Chub-E. This project features a user-friendly Tkinter GUI, precise temperature conversions for PRT and Type S thermocouples, and seamless data visualization and storage. Perfect for metrology labs and temperature measurement enthusiasts! 🚀

# Features
1. Real-Time Data Acquisition: Connects to the Fluke 1529 via serial communication to capture resistance and EMF data.
2. Accurate Temperature Conversion:
3. Converts PRT resistance using the ITS-90 standard.
4. Converts Type S thermocouple EMF with polynomial coefficients across multiple ranges.

# Interactive Visualization:
1. Real-time plots with Matplotlib, supporting raw (Ohms/mV) or temperature (°C) views.
2. Individual channel plots or combined views in separate windows.
3. Automated Data Logging: Saves timestamped data to Excel files with customizable intervals.
4. Flexible Configuration: Adjust COM port, baud rate, measurement period, and channel units (Ohms or mV) via an intuitive GUI.
5. Time Calibration: Syncs device time with the system clock for accurate timestamps.

# Requirements
Python 3.x
## Libraries:
1. tkinter (GUI framework)
2. pyserial (serial communication)
3. pandas (data handling)
4. matplotlib (plotting)
5. openpyxl (Excel output)



## Install dependencies:

pip install pyserial pandas matplotlib openpyxl

# Getting Started
1. Connect the Device: Plug the Fluke 1529 Chub-E into a COM port.
2. Run the Application: script.py

# Configure Settings:
1. Select the COM port and baud rate (default: 9600).
2. Choose the measurement period (e.g., 1s, 5s, 1min).
3. Set channel units (Ohms for PRT, mV for thermocouples).
4. Start Logging: Click "Start Logging" to begin data collection and visualization.
5. View & Save Data: Monitor real-time plots in the main window or open separate channel windows.
Data is automatically saved to Excel files (fluke_1529_YYYYMMDD.xlsx) on the desktop.



# Screenshots
### Coming soon: GUI screenshots showcasing real-time plots and controls!

# Notes

1. Ensure the Fluke 1529 is properly connected and powered on before starting.
2. Data is saved every 60 records or 300 seconds to avoid excessive file writes.
3. The application supports up to four channels, each configurable for PRT or thermocouple measurements.

# Contributing
Contributions are welcome! Feel free to open issues or submit pull requests to enhance functionality or fix bugs. 
Ideas for new features:
1. Support for additional thermocouple types.
2. Enhanced plot customization options.
3. Web-based interface integration.

# License
This project is licensed under the MIT License. See the LICENSE file for details.

Developed with precision for the Department of Temperature & Humidity Metrology(1.04) at CSIR - National Physical Laboratory.
