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
    if not numbers_str:
        return "Nenhum número enviado.", 400

    try:
        numbers = [int(x.strip()) for x in numbers_str.split(",")]
    except ValueError:
        return "Input inválido, por favor enviar apenas números separados por vírgulas.", 400

    # Create Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Numbers"

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
