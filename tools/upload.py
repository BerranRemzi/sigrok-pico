import sys
import os
import shutil
import serial
import serial.tools.list_ports
import win32api
import time

# Check if command line arguments are provided
if len(sys.argv) < 2:
    print("Usage: python script.py <filename>")
    sys.exit(1)

# Get the file name from the first command line argument
file_name = sys.argv[1]

# Verify the file exists
if not os.path.isfile(file_name):
    print(f"Error: File '{file_name}' does not exist.")
    sys.exit(1)

# Function to open and close all available serial ports with a specified baud rate
def switch_to_bootloader(baud_rate=1200):
    returnValue = False
    ports = serial.tools.list_ports.comports()
    if not ports:
        print("No serial ports found.")
        return False
    
    
    for port in ports:
        try:
            ser = serial.Serial(port.device, baud_rate, timeout=1)
            #time.sleep(1)  # Wait for 1 second after opening the port
            ser.close()  # Close the port after opening
            print(f"Forcing reset using {baud_rate}bps open/close on port {port.device}")
            returnValue = True
        except Exception as e:
            continue
    return returnValue

# Function to find drive with volume name "RPI-RP2"
def find_rpi_drive():
    print("Looking for upload drive...")
    time.sleep(1)  # Wait for 1 second
    # Get all available drives
    drives = win32api.GetLogicalDriveStrings().split('\x00')[:-1]
    for drive in drives:
        try:
            volume_info = win32api.GetVolumeInformation(drive)
            if volume_info[0] == "RPI-RP2":
                return drive
        except Exception as e:
            # Skip drives that can't be accessed
            print(f"Could not access drive {drive}: {e}")
    return None

# First, try to handle serial ports
if switch_to_bootloader():

    # Find the RPI-RP2 drive after handling serial ports
    rpi_drive = find_rpi_drive()

    if not rpi_drive:
        print("Error: Drive with volume name 'RPI-RP2' not found.")
        sys.exit(1)
    else:
        print(f"Auto-detected: {rpi_drive}")

    # Copy the file to the RPI-RP2 drive
    destination = os.path.join(rpi_drive, os.path.basename(file_name))
    print(f"Uploading {file_name}")
    try:
        shutil.copy(file_name, destination)
        print(f"Loading into Flash: [==============================]  100%")
    except Exception as e:
        print(f"Failed to copy file: {e}")
else:
    print("Serial port handling failed; file copy operation aborted.")
