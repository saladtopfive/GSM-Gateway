#!/bin/bash
set -e

# Flask musi być PID 1 procesu
exec /usr/bin/python3 app.py
