#!/usr/bin/env python3
import serial
import time
import datetime
import pytz
import os
import logging
import sys
import re
from openpyxl import load_workbook

# ===================== CONFIG =====================

PORT = "/dev/serial/by-id/usb-SimTech__Incorporated_SimTech__Incorporated_0123456789ABCDEF-if04-port0"
BAUD = 115200
CHECK_INTERVAL = 60
LOCAL_TZ = pytz.timezone("Europe/Warsaw")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX_FILE = os.path.join(SCRIPT_DIR, "schedule.xlsx")

PHONE_REGEX = re.compile(r"^\+48\d{9}$")

# ===================== LOGGING =====================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    stream=sys.stdout,
)

log = logging.getLogger("gsm-forwarder")

# ===================== GSM =====================

def open_modem():
    try:
        ser = serial.Serial(PORT, BAUD, timeout=2)
        log.info("Połączono z modemem: %s", PORT)
        return ser
    except Exception as e:
        log.error("Nie można otworzyć modemu: %s", e)
        return None


def send_at(ser, cmd, delay=0.5):
    ser.write((cmd + "\r").encode())
    time.sleep(delay)
    resp = ser.read_all().decode(errors="ignore").strip()
    log.info("AT> %s | %s", cmd, resp)
    return resp


def enable_forwarding(ser, number):
    log.info("Ustawiam przekierowanie na %s", number)
    send_at(ser, f'AT+CCFC=0,3,"{number}",145')


def disable_forwarding(ser):
    log.info("Wyłączam przekierowanie")
    send_at(ser, "AT+CCFC=0,0")

# ===================== XLSX =====================

def load_schedule():
    if not os.path.exists(XLSX_FILE):
        log.error("Brak pliku schedule.xlsx")
        return []

    wb = load_workbook(XLSX_FILE, data_only=True)

    if "schedule" not in wb.sheetnames:
        log.error("Brak arkusza 'schedule'")
        return []

    ws = wb["schedule"]
    schedule = []

    for row_idx in range(2, ws.max_row + 1):
        row = [
            ws[f"A{row_idx}"].value,
            ws[f"B{row_idx}"].value,
            ws[f"C{row_idx}"].value,
            ws[f"D{row_idx}"].value,
            ws[f"E{row_idx}"].value,
            ws[f"F{row_idx}"].value,
        ]

        if all(v is None for v in row):
            continue

        if any(v is None for v in row[:6]):
            log.error("Wiersz %d: niekompletne dane – pomijam", row_idx)
            continue

        start_date, end_date, start_time, end_time, name, number = row

        number = str(number).strip()

        if not PHONE_REGEX.match(number):
            log.error("❌ Odrzucono nieprawidłowy numer: %s", number)
            log.error("Wiersz %d: numer odrzucony", row_idx)
            continue

        try:
            start_dt = LOCAL_TZ.localize(
                datetime.datetime.combine(start_date, start_time)
            )
            end_dt = LOCAL_TZ.localize(
                datetime.datetime.combine(end_date, end_time)
            )
        except Exception as e:
            log.error("Wiersz %d: błąd daty/czasu: %s", row_idx, e)
            continue

        schedule.append((start_dt, end_dt, number, name))

    return schedule


def find_active(schedule):
    now = datetime.datetime.now(LOCAL_TZ)
    next_entry = None

    for start, end, number, name in sorted(schedule, key=lambda x: x[0]):
        if start <= now <= end:
            return ("current", start, end, number, name)
        if start > now and next_entry is None:
            next_entry = ("next", start, end, number, name)

    return next_entry

# ===================== MAIN =====================

def main():
    log.info("Start GSM Forwarder")

    ser = None
    last_number = None

    while True:
        if ser is None or not ser.is_open:
            ser = open_modem()
            if ser:
                send_at(ser, "AT")
                send_at(ser, "ATE0")
                send_at(ser, "AT+CLIP=1")
            else:
                time.sleep(5)
                continue

        schedule = load_schedule()
        result = find_active(schedule)

        now_str = datetime.datetime.now(LOCAL_TZ).strftime("%Y-%m-%d %H:%M:%S")

        try:
            if result and result[0] == "current":
                _, start, end, number, name = result
                if number != last_number:
                    log.info(
                        "%s | Aktywne przekierowanie: %s (%s → %s)",
                        now_str, name, start, end
                    )
                    enable_forwarding(ser, number)
                    last_number = number
                else:
                    log.info("%s | Aktualne przekierowanie: %s", now_str, name)

            else:
                if last_number is not None:
                    log.info("%s | Brak aktywnego przedziału – wyłączam", now_str)
                    disable_forwarding(ser)
                    last_number = None
                else:
                    log.info("%s | Aktualne przekierowanie: Brak", now_str)

        except Exception as e:
            log.error("Błąd komunikacji z modemem: %s", e)
            try:
                ser.close()
            except Exception:
                pass
            ser = None

        time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    main()
