import os
import io
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


# ── Helpers ──────────────────────────────────────────────────────────────
def read_excel_from_bytes(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or r[0] is None:
            continue
        voucher_raw, hotel, room_type, rate = r[0], r[1], r[2], r[3]
        if hotel is None or rate is None:
            continue
        rows.append({
            "voucher": clean_voucher_no(voucher_raw),
            "hotel": str(hotel).strip(),
            "room_type": str(room_type).strip(),
            "rate": float(rate),
        })
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
    except Exception as e:
        return jsonify({"error": f"Failed to read Excel: {str(e)}"}), 400

    if not rows:
        return jsonify({"error": "No valid rows found in Excel"}), 400

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
