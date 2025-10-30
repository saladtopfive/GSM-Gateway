#!/usr/bin/env python3
import serial
import time
from datetime import datetime
import re
import csv

# ==== KONFIGURACJA ====
PORT = "/dev/ttyUSB2"          # port modemu
BAUD = 115200                   # prędkość
FORWARD_NUMBER = "+48500475472" # numer, na który przekierowujemy
LOG_FILE = "call_log.csv"       # plik logu

# ==== FUNKCJE ====
def send_command(ser, cmd, delay=1):
    """Wysyła komendę AT i zwraca odpowiedź modemu"""
    ser.write((cmd + "\r").encode())
    time.sleep(delay)
    resp = ser.read_all().decode(errors="ignore").strip()
    print(f"> {cmd}\n{resp}\n")
    return resp

def set_forward(ser):
    """Ustawienie przekierowania bezwarunkowego (voice)"""
    return send_command(ser, f'AT+CCFC=0,3,"{FORWARD_NUMBER}",145')

def check_status(ser):
    """Sprawdzenie statusu przekierowania"""
    return send_command(ser, 'AT+CCFC=0,2')

def disable_forward(ser):
    """Wyłączenie przekierowania"""
    return send_command(ser, 'AT+CCFC=0,0')

def log_call(number):
    """Zapisuje numer i timestamp do pliku CSV"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([timestamp, number])
    print(f"💾 Zapisano połączenie: {number} o {timestamp}")

def parse_clip(line):
    """Wyciąga numer z +CLIP"""
    match = re.search(r'\+CLIP: "(\+?\d+)"', line)
    if match:
        return match.group(1)
    return None

# ==== SKRYPT ====
def main():
    print("🔌 Łączenie z modemem...")
    ser = serial.Serial(PORT, BAUD, timeout=1)

    # Podstawowa konfiguracja modemu
    send_command(ser, "AT", 0.5)
    send_command(ser, "ATE0", 0.2)     # echo off
    send_command(ser, "AT+CLIP=1", 0.2) # pokazywanie caller ID

    # Ustaw przekierowanie
    print("📡 Ustawiam przekierowanie bezwarunkowe...")
    set_forward(ser)

    # Sprawdzenie statusu
    print("📡 Status przekierowania:")
    check_status(ser)

    # Pętla główna z logowaniem połączeń
    print("📡 Nasłuchiwanie połączeń przychodzących...")
    try:
        buffer = ""
        while True:
            data = ser.read(ser.in_waiting or 1).decode(errors="ignore")
            buffer += data
            if "\n" in buffer:
                lines = buffer.split("\n")
                for line in lines[:-1]:
                    line = line.strip()
                    if line:
                        print(line)
                        number = parse_clip(line)
                        if number:
                            log_call(number)
                buffer = lines[-1]
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("⛔ Wyłączanie przekierowania...")
        disable_forward(ser)
        ser.close()
        print("✅ Zakończono.")

if __name__ == "__main__":
    main()
