# Fluke-1529-Data-Logger-Temperature-conversion

A Python-based data logging application for the Fluke 1529 Chub-E temperature measurement system. This project provides real-time temperature monitoring and conversion for PRT (resistance) and Type S thermocouple (EMF) sensors, featuring a Tkinter GUI with live plotting, data logging to Excel, and SCPI command integration for serial communication.

# Features
1. Real-Time Data Acquisition: Reads temperature and resistance data via serial communication from the Fluke 1529.
2. Temperature Conversion: Converts PRT resistance using ITS-90 and Type S thermocouple EMF using polynomial coefficients.
3. Interactive GUI: Displays real-time values, plots data (raw or temperature), and supports separate channel windows using Matplotlib.
4. Data Logging: Automatically saves data to Excel files with customizable save intervals.

# Configuration: Supports channel unit selection (Ohms or mV) and measurement-period adjustments.
1. Requirements: Python 3.x
2. Libraries: tkinter, pyserial, pandas, matplotlib, openpyxl

# Usage
1. Connect the Fluke 1529 to a COM port.
2. Run the script and select the COM port, baud rate, and measurement settings.
3. Start logging to view real-time plots and save data to Excel files on the desktop.
