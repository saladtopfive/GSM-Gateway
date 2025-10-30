#!/usr/bin/env python3
import serial
import time
from datetime import datetime
import re
import csv

# ==== KONFIGURACJA ====
PORT = "/dev/ttyUSB2"
BAUD = 115200
FORWARD_NUMBER = "+48782335253"
LOG_FILE = "call_log.csv"

# ==== FUNKCJE ====
def send_command(ser, cmd, delay=1):
    ser.write((cmd + "\r").encode())
    time.sleep(delay)
    resp = ser.read_all().decode(errors="ignore").strip()
    print(f"> {cmd}\n{resp}\n")
    return resp

def set_forward(ser):
    """Ustawienie przekierowania bezwarunkowego (voice)"""
    resp = send_command(ser, f'AT+CCFC=0,3,"{FORWARD_NUMBER}",145')
    print("📡 Przekierowanie ustawione.\n")
    return resp

def check_status(ser):
    """Sprawdzenie statusu przekierowania"""
    resp = send_command(ser, 'AT+CCFC=0,2')
    print("📡 Status przekierowania:\n", resp)
    return resp

def disable_forward(ser):
    """Wyłączenie przekierowania"""
    send_command(ser, 'AT+CCFC=0,0')
    print("📡 Przekierowanie wyłączone.")

def log_call(number, call_type="INCOMING"):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([timestamp, call_type, number])
    print(f"💾 Zapisano połączenie ({call_type}): {number} o {timestamp}")

def parse_clip(line):
    match = re.search(r'\+CLIP: "([^"]+)"', line)
    if match:
        return match.group(1)
    return None

def parse_missed(line):
    match = re.search(r'MISSED_CALL: .* (\d+)', line)
    if match:
        return match.group(1)
    return None

# ==== SKRYPT GŁÓWNY ====
def main():
    print("🔌 Łączenie z modemem...")
    ser = serial.Serial(PORT, BAUD, timeout=1)

    # Podstawowa konfiguracja modemu
    send_command(ser, "AT", 0.5)
    send_command(ser, "ATE0", 0.2)
    send_command(ser, "AT+CLIP=1", 0.2)

    # Ustaw przekierowanie
    set_forward(ser)
    check_status(ser)

    print("📡 Nasłuchiwanie połączeń przychodzących... (Ctrl+C aby przerwać)")

    try:
        buffer = ""
        while True:
            data = ser.read(ser.in_waiting or 1).decode(errors="ignore")
            buffer += data
            if "\n" in buffer:
                lines = buffer.split("\n")
                for line in lines[:-1]:
                    line = line.strip()
                    if not line:
                        continue

                    # Wyświetl wszystko z modemu
                    print("📞 Modem:", line)

                    # Wyłapywanie numerów
                    number = parse_clip(line)
                    call_type = "INCOMING"
                    if not number:
                        number = parse_missed(line)
                        call_type = "MISSED"

                    if number:
                        log_call(number, call_type)

                buffer = lines[-1]
            time.sleep(0.1)

    except KeyboardInterrupt:
        print("⛔ Kończę nasłuchiwanie...")
        disable_forward(ser)
        ser.close()
        print("✅ Zakończono.")

if __name__ == "__main__":
    main()
