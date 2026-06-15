from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, send_file

from booking_logic import build_result, save_upload
from booking_logic.utils import BOOKING_RESULT_FILE, RESULT_FILE, STOCK_RESULT_FILE

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    stock_file = request.files.get("stock_file")
    booking_file = request.files.get("booking_file")
    ref_date_value = request.form.get("ref_date")

    if not stock_file or not booking_file or not ref_date_value:
        return render_template(
            "index.html",
            error="Vui lòng chọn đủ 2 file và nhập ngày tham chiếu.",
        )

    try:
        ref_date = datetime.strptime(ref_date_value, "%Y-%m-%d")

        stock_path = save_upload(stock_file, "stock")
        booking_path = save_upload(booking_file, "booking")
        build_result(stock_path, booking_path, ref_date)        
    except Exception as exc:
        return render_template("index.html", error=f"Lỗi xử lý: {exc}")

    return render_template(
        "index.html",
        success=True,
    )


@app.route("/download")
def download():
    return send_file(RESULT_FILE, as_attachment=True, download_name="result.txt")


@app.route("/download-booking")
def download_booking():
    return send_file(BOOKING_RESULT_FILE, as_attachment=True, download_name="Booking_result.xlsx")

@app.route("/download-stock")
def download_stock():
    return send_file(STOCK_RESULT_FILE, as_attachment=True, download_name="Stock_result.xlsx")


if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)
