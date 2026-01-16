from flask import Flask, request, render_template, jsonify, send_file
from openpyxl import load_workbook
from datetime import datetime
import pytz
import os
import shutil
import tempfile

app = Flask(__name__)

LOCAL_TZ = pytz.timezone("Europe/Warsaw")
UPLOAD_PATH = "/home/kp_rpi_user/GSM-Gateway/stable-release/schedule.xlsx"
ALLOWED_EXTENSIONS = {"xlsx"}


# ===== HELPERS =====

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def load_schedule():
    wb = load_workbook(UPLOAD_PATH)
    ws = wb.active
    rows = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not all(row[:5]):
            continue

        sd, ed, st, et, num = row
        start = LOCAL_TZ.localize(datetime.combine(sd, st))
        end = LOCAL_TZ.localize(datetime.combine(ed, et))
        rows.append((start, end, str(num)))

    return sorted(rows, key=lambda x: x[0])


# ===== ROUTES =====

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/status")
def status():
    now = datetime.now(LOCAL_TZ)
    schedule = load_schedule()

    current = None
    next_one = None

    for start, end, num in schedule:
        if start <= now <= end:
            current = (num, start, end)
        elif start > now and next_one is None:
            next_one = (num, start, end)

    def fmt(entry):
        if not entry:
            return None
        return {
            "number": entry[0],
            "start": entry[1].strftime("%Y-%m-%d %H:%M"),
            "end": entry[2].strftime("%Y-%m-%d %H:%M")
        }

    return jsonify({
        "current": fmt(current),
        "next": fmt(next_one)
    })


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "Brak pliku w żądaniu"}), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({"error": "Nie wybrano pliku"}), 400

    if not allowed_file(file.filename):
        return jsonify({
            "error": "Dozwolony jest wyłącznie plik Excel (.xlsx)"
        }), 400

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        file.save(tmp.name)
        tmp_path = tmp.name

    try:
        wb = load_workbook(tmp_path)
        ws = wb.active

        headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        expected = ["start_date", "end_date", "start_time", "end_time", "number"]

        if headers != expected:
            os.unlink(tmp_path)
            return jsonify({
                "error": "Nieprawidłowa struktura pliku Excel"
            }), 400

    except Exception:
        os.unlink(tmp_path)
        return jsonify({
            "error": "Nie udało się odczytać pliku Excel"
        }), 400

    shutil.move(tmp_path, UPLOAD_PATH)

    return jsonify({"status": "ok"})


@app.route("/download")
def download():
    if not os.path.exists(UPLOAD_PATH):
        return jsonify({"error": "Brak pliku harmonogramu"}), 404

    return send_file(
        UPLOAD_PATH,
        as_attachment=True,
        download_name="schedule.xlsx"
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
