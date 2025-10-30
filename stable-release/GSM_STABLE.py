#!/usr/bin/env python3
import serial
import time

# ==== KONFIGURACJA ====
PORT = "/dev/ttyUSB2"          # port modemu
BAUD = 115200                   # prędkość
FORWARD_NUMBER = "+48782335253" # numer, na który przekierowujemy

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

# ==== SKRYPT ====
def main():
    print("🔌 Łączenie z modemem...")
    ser = serial.Serial(PORT, BAUD, timeout=2)
    
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

    # Program pozostaje w pętli, Ctrl+C do wyłączenia
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("⛔ Wyłączanie przekierowania...")
        disable_forward(ser)
        ser.close()
        print("✅ Zakończono.")

if __name__ == "__main__":
    main()

