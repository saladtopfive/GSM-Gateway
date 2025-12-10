#!/usr/bin/env python3
import serial
import time
import datetime
from openpyxl import load_workbook
import pytz
import os

# ==== CONFIG ====
PORT = "/dev/ttyUSB2"
BAUD = 115200
XLSX_FILE = "schedule.xlsx"
LOCAL_TZ = pytz.timezone("Europe/Warsaw")

# ==== HELP FUNCTIONS  ====
def send_command(ser, cmd, delay=1):
    """Wysyła komendę AT i zwraca odpowiedź modemu"""
    ser.write((cmd + "\r").encode())
    time.sleep(delay)
    resp = ser.read_all().decode(errors="ignore").strip()
    print(f"> {cmd}\n{resp}\n")
    return resp

def set_forward(ser, number):
    """Ustawienie przekierowania bezwarunkowego (voice)"""
    return send_command(ser, f'AT+CCFC=0,3,"{number}",145')

def disable_forward(ser):
    """Wyłączenie przekierowania"""
    return send_command(ser, 'AT+CCFC=0,0')

def read_schedule(path):
    """Wczytuje plik XLSX i zwraca listę rekordów"""
    if not os.path.exists(path):
        print(f"❌ Plik {path} nie istnieje.")
        return []

    try:
        wb = load_workbook(path)
        ws = wb.active
        rows = []
        for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not all(row[:5]):
                continue
            start_date, end_date, start_time, end_time, number = row
            try:
                start_dt = datetime.datetime.combine(start_date, start_time)
                end_dt = datetime.datetime.combine(end_date, end_time)
                start_dt = LOCAL_TZ.localize(start_dt)
                end_dt = LOCAL_TZ.localize(end_dt)
                rows.append((start_dt, end_dt, str(number)))
            except Exception as e:
                print(f"⚠️ Błąd w wierszu {i}: {e}")
        return rows
    except Exception as e:
        print(f"❌ Błąd odczytu XLSX: {e}")
        return []

def find_active_forward(schedule):
    """Zwraca numer przekierowania, jeśli obecny czas mieści się w którymś przedziale"""
    now = datetime.datetime.now(LOCAL_TZ)
    for start, end, number in schedule:
        if start <= now <= end:
            return number, start, end
    return None, None, None

# ==== MAIN ====
def main():
    print("🔌 Łączenie z modemem...")
    ser = serial.Serial(PORT, BAUD, timeout=2)

    send_command(ser, "AT", 0.5)
    send_command(ser, "ATE0", 0.2)
    send_command(ser, "AT+CLIP=1", 0.2)

    last_number = None

    try:
        while True:
            schedule = read_schedule(XLSX_FILE)
            number, start, end = find_active_forward(schedule)
            now = datetime.datetime.now(LOCAL_TZ).strftime("%Y-%m-%d %H:%M:%S")

            if number and number != last_number:
                print(f"🕓 {now} | Aktywne przekierowanie: {number} ({start} → {end})")
                set_forward(ser, number)
                last_number = number
            elif not number and last_number is not None:
                print(f"🕓 {now} | Brak aktywnego przedziału — wyłączam przekierowanie.")
                disable_forward(ser)
                last_number = None
            else:
                status = f"{number}" if number else "Brak"
                print(f"🕓 {now} | Aktualne przekierowanie: {status}")

            time.sleep(60)

    except KeyboardInterrupt:
        print("⛔ Wyłączam przekierowanie...")
        disable_forward(ser)
        ser.close()
        print("✅ Zakończono.")

if __name__ == "__main__":
    main()
