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
CHECK_INTERVAL = 60  # seconds
LOCAL_TZ = pytz.timezone("Europe/Warsaw")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX_FILE = os.path.join(SCRIPT_DIR, "schedule.xlsx")

# polski numer: +48XXXXXXXXX albo XXXXXXXXX
PHONE_REGEX = re.compile(r"^(?:\+48)?\d{9}$")

# ===================== LOGGING =====================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    stream=sys.stdout,
)

log = logging.getLogger("gsm-forwarder")

# ===================== GSM HELPERS =====================

def open_modem():
    try:
        ser = serial.Serial(PORT, BAUD, timeout=2)
        log.info("Połączono z modemem: %s", PORT)
        return ser
    except Exception as e:
        log.error("Nie można otworzyć portu %s: %s", PORT, e)
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

# ===================== XLSX HELPERS =====================

def normalize_number(value):
    if value is None:
        return None

    number = str(value).strip().replace(" ", "")
    if number.startswith(("'", "’", "‘")):
        number = number[1:]

    if PHONE_REGEX.fullmatch(number):
        return number

    log.error("❌ Odrzucono nieprawidłowy numer: %s", value)
    return None


def load_contacts(wb):
    phonebook = {}

    if "contacts" not in wb.sheetnames:
        log.warning("Brak arkusza 'contacts'")
        return phonebook

    ws = wb["contacts"]

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 2:
            continue

        name, number = row[0], row[1]
        if not name or not number:
            continue

        phonebook[str(name).strip()] = str(number).strip()

    return phonebook


def load_schedule():
    if not os.path.exists(XLSX_FILE):
        log.warning("Brak pliku XLSX: %s", XLSX_FILE)
        return []

    wb = load_workbook(XLSX_FILE, data_only=True)
    log.info("Dostępne arkusze XLSX: %s", wb.sheetnames)

    schedule_ws = wb[wb.sheetnames[0]]
    phonebook = load_contacts(wb)

    schedule = []

    for idx, row in enumerate(
        schedule_ws.iter_rows(min_row=2, values_only=True), start=2
    ):
        if not row or len(row) < 6:
            continue

        start_date, end_date, start_time, end_time, name, raw_number = row[:6]

        if not all([start_date, end_date, start_time, end_time]):
            continue

        # numer z kolumny albo z contacts
        number = normalize_number(raw_number)
        if not number and name in phonebook:
            number = normalize_number(phonebook[name])

        if not number:
            log.error("Wiersz %d: numer odrzucony", idx)
            continue

        try:
            start_dt = LOCAL_TZ.localize(
                datetime.datetime.combine(start_date, start_time)
            )
            end_dt = LOCAL_TZ.localize(
                datetime.datetime.combine(end_date, end_time)
            )
        except Exception as e:
            log.error("Wiersz %d: błąd daty/czasu: %s", idx, e)
            continue

        schedule.append((start_dt, end_dt, number))

    return schedule


def find_active_forward(schedule):
    now = datetime.datetime.now(LOCAL_TZ)
    for start, end, number in schedule:
        if start <= now <= end:
            return number, start, end
    return None, None, None

# ===================== MAIN LOOP =====================

def main():
    log.info("Start GSM Forwarder")

    ser = None
    last_number = None

    while True:
        # --- modem ---
        if ser is None or not ser.is_open:
            ser = open_modem()
            if ser:
                send_at(ser, "AT")
                send_at(ser, "ATE0")
                send_at(ser, "AT+CLIP=1")
            else:
                time.sleep(5)
                continue

        # --- schedule ---
        schedule = load_schedule()
        number, start, end = find_active_forward(schedule)

        now_str = datetime.datetime.now(LOCAL_TZ).strftime("%Y-%m-%d %H:%M:%S")

        try:
            if number and number != last_number:
                log.info(
                    "%s | Aktywne przekierowanie: %s (%s → %s)",
                    now_str, number, start, end
                )
                enable_forwarding(ser, number)
                last_number = number

            elif not number and last_number is not None:
                log.info("%s | Brak aktywnego przedziału – wyłączam", now_str)
                disable_forwarding(ser)
                last_number = None

            else:
                log.info(
                    "%s | Aktualne przekierowanie: %s",
                    now_str, number if number else "Brak"
                )

        except Exception as e:
            log.error("Błąd komunikacji z modemem: %s", e)
            try:
                ser.close()
            except Exception:
                pass
            ser = None

        time.sleep(CHECK_INTERVAL)

# ===================== ENTRY =====================

if __name__ == "__main__":
    main()
