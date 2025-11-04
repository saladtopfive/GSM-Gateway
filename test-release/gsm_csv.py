#!/usr/bin/env python3
import serial
import time
import csv
from datetime import datetime

# ==== KONFIGURACJA ====
PORT = "/dev/ttyUSB2"
BAUD = 115200
SCHEDULE_FILE = "schedule.csv"

# ==== FUNKCJE AT ====
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

def check_status(ser):
    """Sprawdzenie statusu przekierowania"""
    return send_command(ser, 'AT+CCFC=0,2')

def disable_forward(ser):
    """Wyłączenie przekierowania"""
    return send_command(ser, 'AT+CCFC=0,0')

# ==== FUNKCJE HARMONOGRAMU ====
def load_schedule():
    """Wczytuje harmonogram z pliku CSV"""
    schedule = []
    with open(SCHEDULE_FILE, newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            schedule.append({
                "day": row["day"].lower(),
                "start": row["start"],
                "end": row["end"],
                "number": row["number"]
            })
    return schedule

def get_current_forward_number(schedule):
    """Zwraca numer przekierowania zgodny z aktualnym dniem i godziną"""
    now = datetime.now()
    day = now.strftime("%a").lower()[:3]  # mon, tue, wed...
    current_time = now.strftime("%H:%M")

    for entry in schedule:
        if entry["day"] == day:
            start = entry["start"]
            end = entry["end"]

            # Obsługa zakresów przez północ
            if start < end:
                if start <= current_time < end:
                    return entry["number"]
            else:
                if current_time >= start or current_time < end:
                    return entry["number"]
    return None

# ==== GŁÓWNY SKRYPT ====
def main():
    print("🔌 Łączenie z modemem...")
    ser = serial.Serial(PORT, BAUD, timeout=2)

    send_command(ser, "AT", 0.5)
    send_command(ser, "ATE0", 0.2)
    send_command(ser, "AT+CLIP=1", 0.2)

    schedule = load_schedule()
    current_number = None

    try:
        while True:
            number = get_current_forward_number(schedule)
            if number and number != current_number:
                print(f"📅 Zmiana numeru przekierowania na {number}")
                set_forward(ser, number)
                current_number = number
            time.sleep(60)  # sprawdzanie co minutę
    except KeyboardInterrupt:
        print("⛔ Wyłączanie przekierowania...")
        disable_forward(ser)
        ser.close()
        print("✅ Zakończono.")

if __name__ == "__main__":
    main()
