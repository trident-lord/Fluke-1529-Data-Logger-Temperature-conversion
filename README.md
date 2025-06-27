🌡️ Fluke 1529 Chub-E Data Logger
Copyright (c) 2025 trident-lord

A sleek and robust Python application designed for real-time temperature monitoring and data logging with the Fluke 1529 Chub-E. This project features a user-friendly Tkinter GUI, precise temperature conversions for PRT and Type S thermocouples, and seamless data visualization and storage. Perfect for metrology labs and temperature measurement enthusiasts! 🚀
✨ Features

📡 Real-Time Data Acquisition: Connects to the Fluke 1529 via serial communication to capture resistance and EMF data.
🌡️ Accurate Temperature Conversion:
Converts PRT resistance using the ITS-90 standard.
Converts Type S thermocouple EMF with polynomial coefficients across multiple ranges.


📊 Interactive Visualization:
Real-time plots with Matplotlib, supporting raw (Ohms/mV) or temperature (°C) views.
Individual channel plots or combined views in separate windows.


💾 Automated Data Logging: Saves timestamped data to Excel files with customizable intervals.
⚙️ Flexible Configuration: Adjust COM port, baud rate, measurement period, and channel units (Ohms or mV) via an intuitive GUI.
🕒 Time Calibration: Syncs device time with the system clock for accurate timestamps.

🛠️ Requirements

Python 3.x
Libraries:
tkinter (GUI framework)
pyserial (serial communication)
pandas (data handling)
matplotlib (plotting)
openpyxl (Excel output)



Install dependencies:
pip install pyserial pandas matplotlib openpyxl

🚀 Getting Started

Connect the Device: Plug the Fluke 1529 Chub-E into a COM port.
Run the Application:python fluke_1529_data_logger.py


Configure Settings:
Select the COM port and baud rate (default: 9600).
Choose the measurement period (e.g., 1s, 5s, 1min).
Set channel units (Ohms for PRT, mV for thermocouples).


Start Logging: Click "Start Logging" to begin data collection and visualization.
View & Save Data:
Monitor real-time plots in the main window or open separate channel windows.
Data is automatically saved to Excel files (fluke_1529_YYYYMMDD.xlsx) on the desktop.



📈 Screenshots
Coming soon: GUI screenshots showcasing real-time plots and controls!
📝 Notes

Ensure the Fluke 1529 is properly connected and powered on before starting.
Data is saved every 60 records or 300 seconds to avoid excessive file writes.
The application supports up to four channels, each configurable for PRT or thermocouple measurements.

🌟 Contributing
Contributions are welcome! Feel free to open issues or submit pull requests to enhance functionality or fix bugs. Ideas for new features:

Support for additional thermocouple types.
Enhanced plot customization options.
Web-based interface integration.

📜 License
This project is licensed under the MIT License. See the LICENSE file for details.

Developed with precision for temperature metrology at CSIR - National Physical Laboratory. 🧪
