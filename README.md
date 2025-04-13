# TPFanWin - ThinkPad Fan Control for Windows

This project implements a Windows service (`TPFanWinService`) to control the fan speed of a Lenovo ThinkPad based on CPU temperature readings, using a configurable fan curve. It directly interacts with the Embedded Controller (EC) via Port I/O, leveraging information derived from the Linux kernel's `thinkpad_acpi` driver. Designed specifically for Windows environments.

## Features

*   Runs as a background Windows Service (`TPFanWinService`).
*   Reads CPU temperature from the EC.
*   Adjusts fan speed based on a user-defined curve in `config.yaml`.
*   Uses `inpoutx64.dll` for low-level hardware port access on Windows.
*   Provides command-line interface for service management.

## Prerequisites

1.  **Windows Operating System:** This tool is designed for Windows.
2.  **Python:** Tested with Python 3.13. Make sure Python is installed and added to your system PATH.
3.  **Dependencies:** Install required packages using pip:
    ```bash
    pip install -r requirements.txt
    ```
    This installs:
    *   **pywin32:** Required for Windows Service integration.
    *   **PyYAML:** Used for reading the `config.yaml` file.
4.  **pywin32 Post-Install:** After installing `pywin32`, you **MUST** run its post-install script **as Administrator**:
    ```bash
    # Adjust the path based on your Python installation location if necessary.
    # Example for Python 3.13 installed in default location:
    python "C:\Program Files\Python\Python313\Scripts\pywin32_postinstall.py" -install
    ```
    Failure to do this often results in errors when trying to install or run the service.
    (See the [official pywin32 README](https://github.com/mhammond/pywin32/blob/main/README.md#installing-via-pip) for more details.)
5.  **`inpoutx64.dll`:** This driver is required for direct hardware port access.
    *   Download the 64-bit version from: [https://www.highrez.co.uk/downloads/inpout32/](https://www.highrez.co.uk/downloads/inpout32/)
    *   Place `inpoutx64.dll` in the **same directory** as `fan_control_logic.py`.
6.  **Administrator Privileges:** Required for specific commands (`install`, `remove`, `start`, `stop`, `restart`, `debug`, `update`). The `status` command does *not* require admin rights. The service itself needs appropriate permissions to run and access hardware.

## Files

*   `fan_control_logic.py`: Main script implementing the service logic and command-line interface.
*   `ec_control.py`: Handles low-level EC communication (temperature reads, fan control).
*   `config.yaml`: Configuration file (sensor, interval, fan curve).
*   `inpoutx64.dll`: Hardware access driver (must be present).
*   `requirements.txt`: Python package dependencies.
*   `README.md`: This file.
*   `LICENSE`: Project license (MIT).

## Configuration (`config.yaml`)

*   `sensor_index`: (0-7) The EC temperature sensor to monitor. Default: 0. Identifying the correct sensor might require experimentation or comparison with other monitoring tools.
*   `update_interval_seconds`: How often (in seconds) the service checks temperature and adjusts the fan. Default: 5.
*   `fan_curve`: Defines temperature thresholds (°C) and corresponding fan levels (0-7 for manual speeds, 0x80 for BIOS/Auto control). **Must be sorted by temperature.**

## Recommended Deployment

For running as a service on a target system:
1. Create a stable directory (e.g., `C:\Program Files\TPFanWin`).
2. Copy `fan_control_logic.py`, `ec_control.py`, `config.yaml`, and `inpoutx64.dll` into this directory.
3. Ensure Python and prerequisites (including pywin32 post-install) are met on the system.
4. Open an **Administrator** terminal, `cd` into the deployment directory, and run `python .\fan_control_logic.py install`.

## Usage

Commands are executed using `python .\fan_control_logic.py <command>` from the directory containing the script.

**Run these commands from a standard Command Prompt or PowerShell:**

*   **Check Service Status:**
    ```bash
    python .\fan_control_logic.py status
    ```
    *(Outputs: Running, Stopped, Not Installed, etc. for `TPFanWinService`)*

**Run these commands from an Administrator Command Prompt or PowerShell:**

*   **Install the Service:**
    ```bash
    python .\fan_control_logic.py install
    ```
*   **Start the Service:**
    ```bash
    python .\fan_control_logic.py start
    # OR: net start TPFanWinService
    ```
*   **Stop the Service:**
    ```bash
    python .\fan_control_logic.py stop
    # OR: net stop TPFanWinService
    ```
*   **Restart the Service:**
    ```bash
    python .\fan_control_logic.py restart
    ```
*   **Remove (Uninstall) the Service:**
    ```bash
    python .\fan_control_logic.py remove
    # OR: sc delete TPFanWinService (after stopping)
    ```
*   **Debug the Service (Run in Console):**
    ```bash
    python .\fan_control_logic.py debug
    ```
    *(Runs the service logic directly in the console instead of as a background service. Press Ctrl+C to stop.)*

## Logging

*   **Interactive Commands:** Output (including errors and status messages from commands like `status`, `install`, etc.) is printed directly to the console (stderr).
*   **Service Operation:** The running `TPFanWinService` logs key events (startup, shutdown, temperature checks, fan changes, errors) to the **Windows Event Viewer**. Look under **Windows Logs -> Application** for events with the source `TPFanWinService`.

## Technical Details & Warnings

*   **EC Communication:** The methods used in `ec_control.py` for interacting with the Embedded Controller (e.g., reading temperatures via port `0x1600`, controlling the fan via port `0x2F`) are based on information reverse-engineered from various sources and analysis of the Linux kernel driver `drivers/platform/x86/thinkpad_acpi.c`. You can find the source here: ([Link to thinkpad_acpi.c on GitHub](https://github.com/torvalds/linux/blob/master/drivers/platform/x86/thinkpad_acpi.c))
*   **`inpoutx64.dll`:** This driver provides direct hardware port access (I/O) on Windows. Running such drivers carries inherent risks. Ensure you download it from the official source.
*   **Risk:** Direct hardware manipulation can be system-specific and potentially unstable. Incorrect commands or configuration *could* lead to unexpected system behavior. Use this software **at your own risk**. Verify your `config.yaml`, especially `sensor_index`, carefully.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgements

*   Inspired by and references information from projects like [thinkfan](https://github.com/vmatare/thinkfan), TPFanControl variants, and the Linux `thinkpad_acpi` driver.
*   Developed with assistance from Cascade by Codeium.
