import sys
import os
import servicemanager # For Windows Service logging
import win32serviceutil # ServiceFramework base class
import win32service     # Service status constants
import win32event       # For service stop signaling
import time
import logging
import yaml
import socket
import winerror # For checking specific error codes

from ec_control import get_temperature, set_fan_level, is_admin, FAN_LEVEL_AUTO

# --- Service Configuration ---
SERVICE_NAME = "TPFanWinService"
SERVICE_DISPLAY_NAME = "TPFanWin Service (ThinkPad Fan Control)"
SERVICE_DESCRIPTION = "Controls ThinkPad fan speed based on temperature using a custom curve."

# Determine base path (where script/config is located)
# When running as a service, the CWD might be system32, so we need the script's dir.
try:
    # This is generally the most reliable way to get the script's directory
    BASE_PATH = os.path.dirname(os.path.abspath(__file__))
except NameError:
    # Fallback if __file__ is not defined (e.g., interactive session, rare cases)
    # sys.argv[0] might be unreliable for services, so os.getcwd() is a last resort.
    BASE_PATH = os.getcwd()

# Construct the full path to the config file
CONFIG_FILE_PATH = os.path.join(BASE_PATH, 'config.yaml')

# --- Default Configuration Values ---
# Used if config.yaml is missing or invalid
DEFAULT_SENSOR_INDEX = 0
DEFAULT_INTERVAL = 5
DEFAULT_FAN_CURVE = [
    (0,  0),
    (50, 1),
    (55, 2),
    (65, 3),
    (75, 5),
    (85, 7),
]

# --- Configuration Loading ---
def load_config(filepath=CONFIG_FILE_PATH):
    """Loads configuration from a YAML file, falling back to defaults."""
    sensor_index = DEFAULT_SENSOR_INDEX
    interval = DEFAULT_INTERVAL
    curve = DEFAULT_FAN_CURVE
    log_func = logging.warning # Default to standard logging
    if not servicemanager.RunningAsService():
        log_func = logging.warning
    else:
        log_func = lambda msg: servicemanager.LogInfoMsg(str(msg)) # Use service logger for warnings here

    try:
        with open(filepath, 'r') as f:
            config = yaml.safe_load(f)

        if not isinstance(config, dict):
            log_func(f"Invalid format in {filepath}. Using default settings.")
            return sensor_index, interval, curve

        loaded_index = config.get('sensor_index', sensor_index)
        if isinstance(loaded_index, int) and 0 <= loaded_index <= 7:
            sensor_index = loaded_index
        else:
            log_func(f"Invalid 'sensor_index' in {filepath}: {loaded_index}. Using default: {sensor_index}")
            sensor_index = DEFAULT_SENSOR_INDEX

        loaded_interval = config.get('update_interval_seconds', interval)
        if isinstance(loaded_interval, (int, float)) and loaded_interval > 0:
            interval = loaded_interval
        else:
            log_func(f"Invalid 'update_interval_seconds' in {filepath}: {loaded_interval}. Using default: {interval}")
            interval = DEFAULT_INTERVAL

        loaded_curve = config.get('fan_curve', curve)
        if isinstance(loaded_curve, list) and all(isinstance(p, list) and len(p) == 2 and isinstance(p[0], int) and isinstance(p[1], int) for p in loaded_curve):
            curve = sorted([tuple(p) for p in loaded_curve])
        else:
             log_func(f"Invalid 'fan_curve' format in {filepath}. Using default curve.")
             curve = DEFAULT_FAN_CURVE

        # Info log only needed if running interactively or debugging service
        if not servicemanager.RunningAsService():
             logging.info(f"Successfully loaded configuration from {filepath}")
        else:
             # When running as service, this might be too verbose for info log
             servicemanager.LogInfoMsg(f"Configuration loaded. Sensor: {sensor_index}, Interval: {interval}s, Curve Points: {len(curve)}")

    except FileNotFoundError:
        log_func(f"{filepath} not found. Using default settings.")
    except yaml.YAMLError as e:
        log_func(f"Error parsing {filepath}: {e}. Using default settings.")
    except Exception as e:
        log_func(f"Unexpected error loading {filepath}: {e}. Using default settings.")

    return sensor_index, interval, curve

# --- Fan Control Logic (from previous version) ---
def get_target_fan_level(temperature, curve):
    """Determines the target fan level based on temperature and the fan curve."""
    if temperature is None:
        return FAN_LEVEL_AUTO # Default to auto if temp reading fails

    target_level = curve[0][1] # Start with lowest level from sorted curve
    for temp_threshold, fan_level in curve:
        if temperature >= temp_threshold:
            target_level = fan_level
        else:
            break
    return target_level

# --- Windows Service Class ---
class TPFanWinService(win32serviceutil.ServiceFramework):
    _svc_name_ = SERVICE_NAME
    _svc_display_name_ = SERVICE_DISPLAY_NAME
    _svc_description_ = SERVICE_DESCRIPTION

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60) # Default timeout for network connections, though not used here
        self.is_running = True
        self.config = None
        self.current_level = FAN_LEVEL_AUTO # Assume auto initially
        self.interval_ms = DEFAULT_INTERVAL * 1000

    def SvcStop(self):
        servicemanager.LogInfoMsg(f"{SERVICE_NAME}: Received stop request.")
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.is_running = False

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                            servicemanager.PYS_SERVICE_STARTED,
                            (self._svc_name_, ''))
        try:
            servicemanager.LogInfoMsg(f"{SERVICE_NAME}: Service starting main loop.")
            self.main()
            servicemanager.LogInfoMsg(f"{SERVICE_NAME}: Service main loop finished.")
            # Service cleanup happens after main returns
        except Exception as e:
            servicemanager.LogErrorMsg(f"{SERVICE_NAME}: Unhandled exception in main loop: {e}")
            # Optionally report SERVICE_STOPPED on unhandled exception
            # self.ReportServiceStatus(win32service.SERVICE_STOPPED)
        finally:
            # Ensure status is reported as stopped when SvcDoRun exits
            # This might already be handled by SvcStop setting the event
            self.ReportServiceStatus(win32service.SERVICE_STOPPED)
            servicemanager.LogInfoMsg(f"{SERVICE_NAME}: Service run loop ended.")

    def main(self):
        """Main service logic loop."""
        servicemanager.LogInfoMsg(f"{SERVICE_NAME}: Loading configuration.")
        try:
            self.config = load_config()
            self.interval_ms = self.config[1] * 1000
            sensor_index = self.config[0]
            fan_curve = self.config[2]
            servicemanager.LogInfoMsg(f"{SERVICE_NAME}: Config loaded. Interval: {self.config[1]}s, Sensor: {sensor_index}")
        except Exception as e:
            servicemanager.LogErrorMsg(f"{SERVICE_NAME}: Failed to load config: {e}. Using defaults.")
            # Fallback or stop service? For now, let it stop.
            self.ReportServiceStatus(win32service.SERVICE_STOPPED)
            return

        while self.is_running:
            # Wait for stop signal or timeout
            rc = win32event.WaitForSingleObject(self.hWaitStop, self.interval_ms)

            # Check if stop was signalled
            if rc == win32event.WAIT_OBJECT_0:
                servicemanager.LogInfoMsg(f"{SERVICE_NAME}: Stop event received, exiting loop.")
                break # Exit loop if service is stopping

            # --- Core Fan Control Logic ---
            try:
                temp_c = get_temperature(sensor_index)
                if temp_c is None:
                    servicemanager.LogWarningMsg(f"{SERVICE_NAME}: Failed to read temperature from sensor {sensor_index}.")
                    continue # Skip this cycle if temp read fails

                target_level = get_target_fan_level(temp_c, fan_curve)

                # Only update fan if the target level has changed
                if self.current_level != target_level:
                    servicemanager.LogInfoMsg(f"{SERVICE_NAME}: Temp {temp_c}°C -> Fan Level {target_level:#04x}")
                    if set_fan_level(target_level):
                        self.current_level = target_level
                        servicemanager.LogInfoMsg(f"{SERVICE_NAME}: Fan level set successfully.")
                    else:
                        servicemanager.LogErrorMsg(f"{SERVICE_NAME}: Failed to set fan level {target_level:#04x}.")
                else:
                    # Optional: Log periodically even if level doesn't change
                    # servicemanager.LogInfoMsg(f"{SERVICE_NAME}: Temp {temp_c}°C, Fan Level {self.current_level:#04x} (unchanged)")
                    pass

            except Exception as e:
                servicemanager.LogErrorMsg(f"{SERVICE_NAME}: Error in fan control cycle: {e}")
                # Consider adding a backoff delay here if errors persist

        # --- End of Loop - Service Stopping --- #
        servicemanager.LogInfoMsg(f"{SERVICE_NAME}: Exiting main loop.")
        # Attempt to set fan back to auto/BIOS control on service stop
        try:
            servicemanager.LogInfoMsg(f"{SERVICE_NAME}: Setting fan to AUTO mode on exit.")
            if set_fan_level(FAN_LEVEL_AUTO):
                servicemanager.LogInfoMsg(f"{SERVICE_NAME}: Fan set to AUTO mode on exit.")
            else:
                servicemanager.LogErrorMsg(f"{SERVICE_NAME}: Failed to set fan to AUTO mode on exit.")
        except Exception as e:
            servicemanager.LogErrorMsg(f"{SERVICE_NAME}: Error setting fan to AUTO on exit: {e}")

# --- Helper Function for Status --- #
def print_service_status():
    """Queries and prints the current status of the service."""
    logger = logging.getLogger(__name__)
    try:
        # QueryServiceStatus returns tuple: (serviceType, serviceState, ...)
        status_info = win32serviceutil.QueryServiceStatus(SERVICE_NAME)
        state = status_info[1] # serviceState is the second element
        status_map = {
            win32service.SERVICE_STOPPED: "Stopped",
            win32service.SERVICE_START_PENDING: "Start Pending",
            win32service.SERVICE_STOP_PENDING: "Stop Pending",
            win32service.SERVICE_RUNNING: "Running",
            win32service.SERVICE_CONTINUE_PENDING: "Continue Pending",
            win32service.SERVICE_PAUSE_PENDING: "Pause Pending",
            win32service.SERVICE_PAUSED: "Paused",
        }
        status_string = status_map.get(state, f"Unknown ({state})")
        print(f"Service Status ({SERVICE_NAME}): {status_string}")
        logger.info(f"Service Status ({SERVICE_NAME}): {status_string}")
    except win32service.error as e:
        if e.winerror == winerror.ERROR_SERVICE_DOES_NOT_EXIST:
            print(f"Service Status ({SERVICE_NAME}): Not Installed")
            logger.info(f"Service Status ({SERVICE_NAME}): Not Installed")
        else:
            print(f"Error querying service status: {e}")
            logger.error(f"Error querying service status: {e}")
    except Exception as e:
        print(f"Unexpected error querying service status: {e}")
        logger.error(f"Unexpected error querying service status: {e}", exc_info=True)


# --- Main Execution (Handles Service Commands) ---
# List of commands known to require Administrator privileges
ADMIN_COMMANDS = ['install', 'remove', 'start', 'stop', 'restart', 'debug', 'update']

if __name__ == "__main__":
    # Basic logging setup for interactive runs or service startup errors
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    # Log to stderr for interactive commands
    logging.basicConfig(level=logging.INFO, format=log_format, stream=sys.stderr)
    logger = logging.getLogger(__name__) # Use a specific logger

    if len(sys.argv) == 1:
        # Called without arguments, try to run as a service instance
        logger.info("No arguments provided, attempting to start service dispatcher.")
        try:
            servicemanager.Initialize()
            servicemanager.PrepareToHostSingle(TPFanWinService)
            servicemanager.StartServiceCtrlDispatcher()
        except win32service.error as details:
            import winerror
            if details.winerror == winerror.ERROR_FAILED_SERVICE_CONTROLLER_CONNECT:
                 # Updated error message for clarity
                 logger.error("Cannot start service from command line.")
                 print("\nError: This script cannot be started directly without arguments.")
                 print("To manage the service, use commands like:")
                 print(f"  python {os.path.basename(sys.argv[0])} install")
                 print(f"  net start {SERVICE_NAME}")
                 print(f"  python {os.path.basename(sys.argv[0])} status")
                 print("  (Run 'install', 'remove', 'start', 'stop' commands as Administrator)")
            else:
                 logger.error(f"Error starting service dispatcher: {details}")
            sys.exit(1) # Exit with error code
    else:
         # Called with arguments (install, start, stop, debug etc)
         command = sys.argv[1].lower()
         logger.info(f"Command received: {command} with args: {sys.argv[1:]}")

         # --- Handle 'status' command separately ---
         if command == 'status':
             print_service_status()
             sys.exit(0) # Exit after printing status

         # --- Handle other commands (install, start, stop, debug etc) ---

         # Check for admin rights if required by the command
         if command in ADMIN_COMMANDS:
             logger.info(f"Command '{command}' requires Administrator privileges. Checking...")
             if not is_admin():
                 logger.error(f"Command '{command}' requires Administrator privileges.")
                 print(f"\nError: Command '{command}' requires Administrator privileges.")
                 print("Please re-run this command from a Command Prompt or PowerShell opened as Administrator.")
                 sys.exit(1) # Exit with error code
             else:
                 logger.info("Administrator privileges detected.")

         # Proceed to call HandleCommandLine for other standard service commands
         try:
             logger.info(f"Passing arguments to HandleCommandLine: {sys.argv}")
             win32serviceutil.HandleCommandLine(TPFanWinService, argv=sys.argv)
             logger.info(f"HandleCommandLine for '{command}' completed.")
         except SystemExit as se:
             # HandleCommandLine might exit, capture non-zero exit codes
             if se.code != 0:
                 logger.warning(f"HandleCommandLine exited with code: {se.code}")
         except Exception as e:
             logger.exception(f"Error during HandleCommandLine execution for command '{command}': {e}")
             sys.exit(1)
