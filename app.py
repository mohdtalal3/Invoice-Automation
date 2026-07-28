import os
import io
import re
import threading
from collections import defaultdict
from datetime import datetime

import requests
import openpyxl
from flask import Flask, request, session, redirect, url_for, render_template, jsonify
from dotenv import load_dotenv

from main import (
    login, process_voucher, clean_voucher_no, BASE,
)

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY")

PIN = os.getenv("PIN")

# ── Global in-memory state (no files – Vercel safe) ──────────────────────
STATE = {
    "excel_rows": [],
    "unique_vouchers": [],
    "processing": False,
    "results": {},
    "logs": [],
    "done": False,
    "filename": "",
}
LOCK = threading.Lock()


REQUIRED_COLUMNS = ["VOUCHER", "HOTEL", "ROOM RATE"]


class ExcelValidationError(Exception):
    pass


def _normalize_header(h):
    return re.sub(r"[^A-Z]", " ", str(h or "").upper()).strip()
    # collapses things like "VOUCHER # " / "ROOM RATE " into "VOUCHER" / "ROOM RATE"


def read_excel_from_bytes(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if not header_row:
        raise ExcelValidationError("The Excel file is empty or has no header row.")

    normalized_headers = [_normalize_header(h) for h in header_row]

    col_index = {}
    missing = []
    for required in REQUIRED_COLUMNS:
        req_norm = _normalize_header(required)
        found_idx = None
        for idx, h in enumerate(normalized_headers):
            if h == req_norm or h.startswith(req_norm):
                found_idx = idx
                break
        if found_idx is None:
            missing.append(required)
        else:
            col_index[required] = found_idx

    if missing:
        raise ExcelValidationError(
            f"Missing or incorrectly named column(s): {', '.join(missing)}. "
            f"Expected columns: {', '.join(REQUIRED_COLUMNS)}. "
            f"Found headers: {', '.join(str(h) for h in header_row if h)}"
        )

    v_idx, h_idx, r_idx = (
        col_index["VOUCHER"], col_index["HOTEL"], col_index["ROOM RATE"]
    )

    rows = []
    for row_num, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not r or r[v_idx] is None:
            continue
        voucher_raw, hotel, rate = r[v_idx], r[h_idx], r[r_idx]
        if hotel is None or rate is None:
            continue
        try:
            rate_val = float(rate)
        except (TypeError, ValueError):
            raise ExcelValidationError(
                f"Row {row_num}: 'Room Rate' value '{rate}' is not a valid number."
            )
        rows.append({
            "voucher": clean_voucher_no(voucher_raw),
            "hotel": str(hotel).strip(),
            "rate": rate_val,
        })

    if not rows:
        raise ExcelValidationError("No valid data rows found in the Excel file.")

    return rows


def add_log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    STATE["logs"].append(f"[{ts}] {msg}")
    if len(STATE["logs"]) > 300:
        STATE["logs"] = STATE["logs"][-300:]


def process_all():
    """Runs in a background thread – survives tab close."""
    try:
        add_log("Starting processing...")
        sess = requests.Session()

        add_log("Logging into CRM...")
        if not login(sess):
            add_log("ERROR: CRM login failed")
            return
        add_log("CRM login successful")

        excel_rows = STATE["excel_rows"]
        grouped = defaultdict(list)
        for r in excel_rows:
            grouped[r["voucher"]].append(r)

        unique_vouchers = sorted(set(r["voucher"] for r in excel_rows))
        add_log(f"Found {len(unique_vouchers)} unique voucher(s): {unique_vouchers}")

        results = {}
        for voucher_no in unique_vouchers:
            process_voucher(sess, voucher_no, grouped[voucher_no], log_fn=add_log, results=results)
            STATE["results"] = dict(results)

        STATE["results"] = results
        add_log("Processing complete!")
        STATE["done"] = True
    except Exception as e:
        add_log(f"ERROR: {str(e)}")
    finally:
        STATE["processing"] = False


# ── Routes ───────────────────────────────────────────────────────────────
@app.route("/", methods=["GET", "POST"])
def login_route():
    if session.get("logged_in"):
        return redirect(url_for("dashboard"))

    error = None
    if request.method == "POST":
        pin = request.form.get("pin", "")
        if pin == PIN:
            session["logged_in"] = True
            return redirect(url_for("dashboard"))
        error = "Invalid PIN. Please try again."

    return render_template("login.html", error=error)


@app.route("/dashboard")
def dashboard():
    if not session.get("logged_in"):
        return redirect(url_for("login_route"))
    return render_template("dashboard.html")


@app.route("/upload", methods=["POST"])
def upload():
    if not session.get("logged_in"):
        return jsonify({"error": "Not logged in"}), 401

    with LOCK:
        if STATE["processing"]:
            return jsonify({"error": "Cannot upload while processing is running"}), 400

    file = request.files.get("excel_file")
    if not file or not file.filename:
        return jsonify({"error": "No file uploaded"}), 400

    file_bytes = file.read()
    try:
        rows = read_excel_from_bytes(file_bytes)
    except ExcelValidationError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Failed to read Excel: {str(e)}"}), 400

    with LOCK:
        STATE["excel_rows"] = rows
        STATE["unique_vouchers"] = sorted(set(r["voucher"] for r in rows))
        STATE["results"] = {}
        STATE["logs"] = []
        STATE["done"] = False
        STATE["filename"] = file.filename

    return jsonify({
        "unique_vouchers": STATE["unique_vouchers"],
        "total_rows": len(rows),
        "filename": STATE["filename"],
    })


@app.route("/start", methods=["POST"])
def start():
    if not session.get("logged_in"):
        return jsonify({"error": "Not logged in"}), 401

    with LOCK:
        if STATE["processing"]:
            return jsonify({"error": "Already processing another file. Please wait."}), 409
        if not STATE["excel_rows"]:
            return jsonify({"error": "No Excel file uploaded"}), 400

        STATE["processing"] = True
        STATE["done"] = False
        STATE["results"] = {}
        STATE["logs"] = []

    thread = threading.Thread(target=process_all, daemon=True)
    thread.start()
    return jsonify({"status": "started"})


@app.route("/status")
def status():
    if not session.get("logged_in"):
        return jsonify({"error": "Not logged in"}), 401

    return jsonify({
        "processing": STATE["processing"],
        "done": STATE["done"],
        "logs": STATE["logs"],
        "results": STATE["results"],
        "unique_vouchers": STATE["unique_vouchers"],
        "filename": STATE["filename"],
        "has_excel": bool(STATE["excel_rows"]),
    })


@app.route("/reset", methods=["POST"])
def reset():
    if not session.get("logged_in"):
        return jsonify({"error": "Not logged in"}), 401
    with LOCK:
        if STATE["processing"]:
            return jsonify({"error": "Cannot reset while processing"}), 400
        STATE["excel_rows"] = []
        STATE["unique_vouchers"] = []
        STATE["results"] = {}
        STATE["logs"] = []
        STATE["done"] = False
        STATE["filename"] = ""
    return jsonify({"status": "ok"})


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login_route"))


if __name__ == "__main__":
    app.run(debug=False, port=5000)
