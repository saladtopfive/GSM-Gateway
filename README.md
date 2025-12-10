# 📞 GSM Call Forwarding Scheduler  
Automatic management of call forwarding on a GSM modem based on a schedule stored in an Excel file.

## Overview
The script periodically reads the `schedule.xlsx` file and checks whether the current time falls within any defined time range.  
If it does — it enables call forwarding to the specified number.  
If not — it disables call forwarding.

Communication with the GSM modem is done via serial port using AT commands.

---

## `schedule.xlsx` structure
The file must contain headers in row 1 and data starting from row 2:

| Start Date | End Date | Start Time | End Time | Number |
|-----------|-----------|------------|----------|--------|
| 2025-01-10 | 2025-01-10 | 08:00 | 16:00 | 123456789 |

Each row defines one active forwarding time window.

---

## Requirements
- Python 3.7+
- Required libraries:
  ```bash
  pip install pyserial openpyxl pytz
  ```
- A GSM modem supporting:
  - `AT+CCFC` (call forwarding)
  - `AT+CLIP`
- Access to `/dev/ttyUSB2` or another configured serial port

---

## Configuration
Editable parameters are located at the top of the script:

```python
PORT = "/dev/ttyUSB2"
BAUD = 115200
XLSX_FILE = "schedule.xlsx"
LOCAL_TZ = pytz.timezone("Europe/Warsaw")
```

---

##  Running the script

```bash
python3 forwarder.py
```

The script will:
- connect to the GSM modem,
- check the schedule every 60 seconds,
- automatically enable or disable call forwarding,
- print all AT commands and modem responses.

---

## Loop logic
Every minute the script performs:

1. Load the schedule from Excel.
2. Determine if the current time falls into any defined time range.
3. If yes — enable call forwarding:
   ```
   AT+CCFC=0,3,"NUMBER",145
   ```
4. If no — disable call forwarding:
   ```
   AT+CCFC=0,0
   ```

The script remembers the last active number to avoid sending redundant commands.

---

## Stopping the script
On Ctrl+C:
- call forwarding is automatically disabled,
- the serial port is closed safely.

---

## Notes
- The Excel file must not be open in edit mode — the script needs to read it.
- All dates and times are interpreted in the `Europe/Warsaw` timezone.
- Invalid rows in the schedule are skipped with a warning.

## License
This project is licensed under the MIT permissive license.
