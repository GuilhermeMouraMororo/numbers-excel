from flask import Flask, render_template, request, send_file, abort
from io import BytesIO
from openpyxl import Workbook
import re
from datetime import datetime

app = Flask(__name__)

def parse_numbers(s: str):
    tokens = re.split(r"[,\s;]+", s.strip())
    nums = []
    for t in tokens:
        if t == "":
            continue
        nums.append(int(t))
    return nums

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    numbers_str = request.form.get("numbers", "").strip()
    n_str = request.form.get("n", "").strip()

    if not numbers_str or not n_str:
        abort(400, "Please provide both the list of numbers and n.")

    try:
        numbers = parse_numbers(numbers_str)
        n = int(n_str)
        if n <= 0:
            raise ValueError
    except ValueError:
        abort(400, "Invalid input. Use integers only and n > 0.")

    wb = Workbook()
    ws = wb.active
    ws.title = "Data"

    for _ in range(n):
        ws.append(numbers)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)

    filename = f"numbers_{datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')}.xlsx"
    return send_file(
        bio,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

if __name__ == "__main__":
