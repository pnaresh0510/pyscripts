"""
This script is for product B Testing. Used to log temperature data from keysight.

Specific channel mapped for testing.

Author: Naresh Penugonda
Version: 1.0  # Initial
Version: 1.01 # rounded off readings to 2 decimal places.
version: 1.02 # Modularity done and added functions to config and scan and initialization

"""
# ------------------- Module import section ---------------------
import pyvisa
import time
from openpyxl import Workbook
from decimal import Decimal, ROUND_HALF_UP

# -------------------- Auto-detect DAQ970A ----------------------
rm = pyvisa.ResourceManager()
daq_address = None
# -------------------- Configuration --------------------
scan_channels = ['112', '101', '102', '103', '104', '105', '116', '117', '118']  # Change to your TC channels
scan_interval = 2  # seconds
scan_cycles = 5  # No. of scan cycles
thermistor_type = 'K'  # K Type thermistor

def daq_init():
    global daq_address
    print("Scanning VISA resources...")
    for resource in rm.list_resources():
        try:
            inst = rm.open_resource(resource)
            idn = inst.query("*IDN?").strip()
            print(f"Found: {idn} at {resource}")
            if "MY58025899" in idn:
                daq_address = resource
                print(f"Selected DAQ970A at {daq_address}")
                inst.close()
                break
            inst.close()
        except Exception as e:
            print(f"Failed to query {resource}: {e}")

        finally:
            print("Done scanning VISA resources...")

    if not daq_address:
        print("No Keysight DAQ970A found.")
        exit()

def daq_cfg():

    # Clear and reset state and identify
    daq.write("ABORt")  # Abort any previous scan safely
    daq.write("*RST;*CLS")
    # daq.write("ROUT:SCAN:DEL:AUTO OFF")  # Avoid auto delays between scans if not needed

    # Configure channels for K-type thermocouples
    for ch in scan_channels:
        daq.write(f"CONF:TEMP TC, {thermistor_type} ,(@{ch})")
        daq.write(f"SENS:TEMP:TRAN:TC:RJUN:TYPE INT,(@{ch})")  # Internal ref junction

    # Setup scan list and triggering
    channel_list = f"@{','.join([str(ch) for ch in scan_channels])}"
    daq.write(f"ROUT:SCAN ({channel_list})")

def daq_scan():
    # Start scan loop
    for i in range(scan_cycles):
        daq.write("INIT")
        time.sleep(10)  # wait for data acquisition

        readings = daq.query("FETC?").strip().split(',')

        # Define precision: 2 decimal places
        precision = Decimal('0.01')

        decimal_data = [str(Decimal(val).quantize(precision, rounding=ROUND_HALF_UP)) for val in readings]

        timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
        row = [timestamp] + decimal_data
        ws.append(row)

        print(f"Scan {i + 1}: {row}")
        time.sleep(scan_interval)

daq_init()                                 # DAQ initialization function call
daq = rm.open_resource(daq_address)
daq.timeout = 10000  # 10 sec timeout
daq_cfg()  # Configuration function call

# Setup Excel logging
wb = Workbook()
ws = wb.active
ws.title = "DAQ Log"
headers = ["Timestamp"] + [f"Ch{ch}" for ch in scan_channels]
ws.append(headers)

daq_scan()                              # daq scanning channels

# Save Excel
file_name = "DAQ970A_Temperature_Log.xlsx"

try:
    wb.save(file_name)
    print(f"Data saved to {file_name}")
except PermissionError:
    pass

# Close instrument
daq.close()
rm.close()
