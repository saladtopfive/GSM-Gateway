#!/usr/bin/env python3
import serial
import time
import datetime
import pytz
import os
import logging
import sys
from openpyxl import load_workbook

# ===================== CONFIG =====================

PORT = "/dev/serial/by-id/usb-SimTech__Incorporated_SimTech__Incorporated_0123456789ABCDEF-if04-port0"
BAUD = 115200
CHECK_INTERVAL = 60  # seconds
LOCAL_TZ = pytz.timezone("Europe/Warsaw")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX_FILE = os.path.join(SCRIPT_DIR, "schedule.xlsx")

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
    """Próbuje otworzyć port modemu (retry-safe)"""
    try:
        ser = serial.Serial(PORT, BAUD, timeout=2)
        log.info("Połączono z modemem: %s", PORT)
        return ser
    except Exception as e:
        log.error("Nie można otworzyć portu %s: %s", PORT, e)
        return None


def send_at(ser, cmd, delay=0.5):
    """Wysyła komendę AT i zwraca odpowiedź"""
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

# ===================== XLSX LOGIC =====================

def load_schedule():
    """Czyta XLSX i zwraca listę (start, end, number)"""
    if not os.path.exists(XLSX_FILE):
        log.warning("Brak pliku XLSX: %s", XLSX_FILE)
        return []

    wb = load_workbook(XLSX_FILE)
    ws = wb.active
    schedule = []

    for row_idx in range(2, ws.max_row + 1):
        values = [ws[f"{c}{row_idx}"].value for c in ("A", "B", "C", "D", "E")]

        if all(v is None for v in values):
            continue

        if any(v is None for v in values):
            log.error("Niekompletny wiersz %d w XLSX – ignoruję", row_idx)
            continue

        start_date, end_date, start_time, end_time, number = values

        try:
            start_dt = LOCAL_TZ.localize(
                datetime.datetime.combine(start_date, start_time)
            )
            end_dt = LOCAL_TZ.localize(
                datetime.datetime.combine(end_date, end_time)
            )
        except Exception as e:
            log.error("Błąd daty/czasu w wierszu %d: %s", row_idx, e)
            continue

        number = str(number).strip()
        if number.startswith(("'", "’", "‘")):
            number = number[1:]

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
                log.info("%s | Brak aktywnego przedziału", now_str)
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
