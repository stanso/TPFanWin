# -*- coding: utf-8 -*-
"""Embedded Controller (EC) communication via inpoutx64.dll for ThinkPads."""

import ctypes
import os
import time
import logging
import sys

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Constants ---

# Standard KCS Interface Ports
EC_DATA_PORT = 0x62
EC_CTRL_PORT = 0x66 # Also Status Port
EC_STATUS_PORT = 0x66 # Also Control Port

# KCS Commands/Status Bits (from Status Port 0x66)
EC_STATUS_OBF = 0x01  # Output Buffer Full
EC_STATUS_IBF = 0x02  # Input Buffer Full
EC_CTRL_READ_CMD = 0x80 # Read command
EC_CTRL_WRITE_CMD = 0x81 # Write command
EC_CMD_WRITE = 0x81

# ThinkPad Specific EC Register Offsets (from Linux kernel thinkpad_acpi.c)
TP_EC_FAN_STATUS = 0x2F      # Read/Write: Fan Level (0-7, 0x40=Full, 0x80=Auto)
TP_EC_FAN_RPM_LSB = 0x84     # Read: Fan Speed RPM Low Byte
TP_EC_FAN_RPM_MSB = 0x85     # Read: Fan Speed RPM High Byte (read after LSB)
TP_EC_FAN_SELECT = 0x31      # Write: Select Fan (0=Fan1, 1=Fan2, if applicable)
TP_EC_TEMP_BASE = 0x78       # Read: Base offset for Temp0-Temp7 (sensors 0-7)
# TP_EC_TEMP_EXT_BASE = 0xC0 # Read: Base offset for Temp8-Temp15 (sensors 8-15, less common)

# Fan Levels
FAN_LEVEL_AUTO = 0x80
FAN_LEVEL_FULL = 0x40
# Levels 0-7 are also valid

# Timing/Timeouts (in seconds)
KCS_WAIT_TIMEOUT = 0.2 # Increased from 0.1
KCS_RETRY_DELAY = 0.001 # 1 ms

# --- Load inpoutx64.dll ---

DLL_NAME = "inpoutx64.dll"
DLL_PATH = os.path.join(os.path.dirname(__file__), DLL_NAME)

_inpoutx64 = None
_Inp32 = None
_Out32 = None

try:
    if not os.path.exists(DLL_PATH):
        raise FileNotFoundError(f"{DLL_NAME} not found at {DLL_PATH}. Ensure it's in the same directory as this script.")
    
    _inpoutx64 = ctypes.windll.LoadLibrary(DLL_PATH)
    
    # Define function prototypes for Port I/O
    # Despite the DLL name, the functions are often exported as Inp32/Out32
    # short _stdcall Inp32(short PortAddress);
    _Inp32 = _inpoutx64.Inp32
    _Inp32.argtypes = [ctypes.c_ushort]
    _Inp32.restype = ctypes.c_ushort

    # void _stdcall Out32(short PortAddress, short Data);
    _Out32 = _inpoutx64.Out32
    _Out32.argtypes = [ctypes.c_ushort, ctypes.c_ushort]
    _Out32.restype = None

    logging.info(f"Successfully loaded {DLL_NAME} and found Inp32/Out32 functions")

except FileNotFoundError as e:
    logging.error(e)
    # Application cannot function without the DLL
    # Consider raising the exception or exiting
    raise
except Exception as e:
    logging.error(f"Failed to load or initialize {DLL_NAME}: {e}")
    # Consider raising the exception or exiting
    raise

# --- Low-Level Port I/O Functions ---

def _read_port(port):
    """Reads a byte from the specified I/O port."""
    if not _Inp32:
        raise RuntimeError(f"{DLL_NAME} not loaded or Inp32 function not found.")
    try:
        return _Inp32(port) & 0xFF # Read byte
    except Exception as e:
        logging.error(f"Error reading from port 0x{port:X}: {e}")
        raise

def _write_port(port, value):
    """Writes a byte to the specified I/O port."""
    if not _Out32:
        raise RuntimeError(f"{DLL_NAME} not loaded or Out32 function not found.")
    try:
        _Out32(port, value & 0xFF) # Write byte
    except Exception as e:
        logging.error(f"Error writing {value:#04x} to port 0x{port:X}: {e}")
        raise

# --- KCS Protocol Implementation ---

def _kcs_wait_ibf_clear():
    """Waits for the KCS Input Buffer Full (IBF) flag to clear."""
    start_time = time.time()
    # retry_count = 0 # No longer needed if not logging retries
    while time.time() - start_time < KCS_WAIT_TIMEOUT:
        status = _read_port(EC_STATUS_PORT)
        if not (status & EC_STATUS_IBF): # IBF is bit 1
            # logging.debug(f"_kcs_wait_ibf_clear: IBF cleared after {retry_count} retries.")
            return True
        # time.sleep(KCS_RETRY_DELAY)
        pass # Use pass for tight loop
        # retry_count += 1

    logging.warning("Timeout waiting for IBF flag to clear.")
    return False

def _kcs_wait_obf_set():
    """Waits for the KCS Output Buffer Full (OBF) flag to set."""
    start_time = time.time()
    # retry_count = 0 # No longer needed if not logging retries
    while time.time() - start_time < KCS_WAIT_TIMEOUT:
        status = _read_port(EC_STATUS_PORT)
        if status & EC_STATUS_OBF: # OBF is bit 0
            # logging.debug(f"_kcs_wait_obf_set: OBF set after {retry_count} retries.")
            return True
        # time.sleep(KCS_RETRY_DELAY)
        pass # Use pass for tight loop
        # retry_count += 1

    logging.warning("Timeout waiting for OBF flag to set.")
    return False

def _read_ec_byte(offset):
    """Reads a byte from the EC at the specified offset using KCS protocol."""
    initial_status = _read_port(EC_CTRL_PORT)
    logging.debug(f"_read_ec_byte({offset:#04x}): Initial EC Status = {initial_status:#04x}")

    if not _kcs_wait_ibf_clear():
        raise TimeoutError("EC KCS timeout: Failed waiting for IBF clear before read command.")

    # Send read command to Control Port
    _write_port(EC_CTRL_PORT, EC_CTRL_READ_CMD)

    if not _kcs_wait_ibf_clear():
        raise TimeoutError("EC KCS timeout: Failed waiting for IBF clear before sending offset.")

    # Send offset to Data Port
    _write_port(EC_DATA_PORT, offset)

    # Step 4: Wait for OBF (Output Buffer Full) to be set
    if not _kcs_wait_obf_set():
        raise TimeoutError("EC KCS timeout: Failed waiting for OBF set after sending offset.")

    # Read data from Data Port
    data = _read_port(EC_DATA_PORT)

    # Step 6: Read status port AFTER reading data to ensure OBF clears
    status_after_read = _read_port(EC_STATUS_PORT)
    logging.debug(f"Read data {hex(data)}, status after read: {hex(status_after_read)}")

    return data

def _write_ec_byte(offset, value):
    """Writes a byte to the specified EC offset using KCS protocol."""
    logging.debug(f"_write_ec_byte({hex(offset)}, {hex(value)}): Initial EC Status = {hex(_read_port(EC_STATUS_PORT))}")

    # Step 1: Wait for IBF (Input Buffer Full) to be clear
    if not _kcs_wait_ibf_clear():
        raise TimeoutError("EC KCS timeout: Failed waiting for IBF clear before sending write cmd.")

    # Step 2: Send Write command (0x81) to Command Port
    _write_port(EC_CTRL_PORT, EC_CMD_WRITE)
    logging.debug("Sent write command (0x81)")

    # Step 3: Wait for IBF to be clear
    if not _kcs_wait_ibf_clear():
        raise TimeoutError("EC KCS timeout: Failed waiting for IBF clear after sending write cmd.")

    # Step 4: Send offset to Data Port
    _write_port(EC_DATA_PORT, offset)
    logging.debug(f"Sent offset {hex(offset)}")

    # Step 5: Wait for IBF to be clear
    if not _kcs_wait_ibf_clear():
        raise TimeoutError("EC KCS timeout: Failed waiting for IBF clear after sending offset.")

    # Step 6: Send data byte to Data Port
    _write_port(EC_DATA_PORT, value)
    logging.debug(f"Sent data {hex(value)}")

    # Step 7: Wait for IBF to be clear (EC has accepted the data)
    if not _kcs_wait_ibf_clear():
        raise TimeoutError("EC KCS timeout: Failed waiting for IBF clear after sending data.")

    logging.debug(f"_write_ec_byte({hex(offset)}, {hex(value)}) completed successfully.")


# --- High-Level Control Functions ---

def get_fan_rpm(fan_index=0):
    """Gets the RPM for the specified fan (default 0). Requires Administrator privileges."""
    # TODO: Add fan selection using TP_EC_FAN_SELECT if fan_index=1 and supported
    if fan_index != 0:
        logging.warning("Multi-fan selection not yet implemented, querying fan 0.")
        # Consider raising NotImplementedError or handling fan selection here
    
    try:
        # IMPORTANT: Read LSB first, then MSB
        rpm_lsb = _read_ec_byte(TP_EC_FAN_RPM_LSB)
        rpm_msb = _read_ec_byte(TP_EC_FAN_RPM_MSB)
        rpm = (rpm_msb << 8) | rpm_lsb
        # Some ECs return 0xFFFF or 0 when fan is off or speed is unavailable
        if rpm == 0xFFFF:
            return 0 # Treat as off
        return rpm
    except (TimeoutError, RuntimeError, OSError) as e:
        logging.error(f"Failed to get fan RPM: {e}")
        return -1 # Indicate error

def get_temperature(sensor_index):
    """Gets the temperature from the specified sensor index (0-7). Requires Administrator privileges."""
    if not 0 <= sensor_index <= 7:
        # Extend range if TP_EC_TEMP_EXT_BASE is used and needed
        raise ValueError("Sensor index must be between 0 and 7.")
    
    offset = TP_EC_TEMP_BASE + sensor_index
    try:
        temp = _read_ec_byte(offset)
        # EC often returns 0x80 (-128 signed) if sensor is disabled or unavailable
        if temp == 0x80:
            logging.warning(f"Temperature sensor {sensor_index} returned 0x80 (possibly disabled/unavailable).")
            return None # Indicate unavailable
        # Temperatures are typically signed bytes
        # Convert unsigned byte read by ctypes to signed byte
        if temp > 127:
            return temp - 256
        else:
            return temp
    except (TimeoutError, RuntimeError, OSError) as e:
        logging.error(f"Failed to get temperature for sensor {sensor_index}: {e}")
        return None # Indicate error

def set_fan_level(level, fan_index=0):
    """Sets the fan control level. Requires Administrator privileges.
    Level can be 0-7, FAN_LEVEL_AUTO (0x80), or FAN_LEVEL_FULL (0x40).
    """
    # TODO: Add fan selection using TP_EC_FAN_SELECT if fan_index=1 and supported
    if fan_index != 0:
        logging.warning("Multi-fan selection not yet implemented, setting fan 0.")
        # Consider raising NotImplementedError or handling fan selection here

    if not (0 <= level <= 7 or level == FAN_LEVEL_AUTO or level == FAN_LEVEL_FULL):
        raise ValueError("Invalid fan level specified.")

    try:
        _write_ec_byte(TP_EC_FAN_STATUS, level)
        logging.info(f"Successfully set fan level to {level:#04x}")
        return True
    except (TimeoutError, RuntimeError, OSError) as e:
        logging.error(f"Failed to set fan level: {e}")
        return False


# --- Helper Functions ---

def is_admin():
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
    except AttributeError:
        # Fallback for environments where IsUserAnAdmin might not be available
        is_admin = False 
        print("WARNING: Could not determine admin status using ctypes.")

    return is_admin

# --- Example Usage (for testing) ---
if __name__ == "__main__":
    # Basic configuration for logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logging.info("EC Control Module loaded.")
    logging.info("This script provides functions for ThinkPad EC interaction.")
    logging.info("Import this module into your application to use its functions.")
