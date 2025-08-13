from flask import Flask, render_template, request, send_file
import openpyxl
from openpyxl import Workbook
import io

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate_excel():
    numbers_str = request.form.get("numbers")
    n_str = request.form.get("n", "").strip()
    if not numbers_str or not n_str:
        return "No numbers provided", 400

    try:
        numbers = [int(x.strip()) for x in numbers_str.split(",")]
    except ValueError:
        return "Invalid input. Please enter only numbers separated by commas.", 400
    try:
        numbers = parse_numbers(numbers_str)
        n = int(n_str)
        if n <= 0:
            raise ValueError
    except ValueError:
        abort(400, "Invalid input. Use integers only and n > 0.")

    # Create Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Numbers"

    # Write numbers to column A
   for _ in range(n):
        ws.append(numbers)

    # Save to memory buffer
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="numbers.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    app.run(debug=True)
