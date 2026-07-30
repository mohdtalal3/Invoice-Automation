import os
import re
import json
import time
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv

load_dotenv()

JOB_URL = "https://paddleocr.aistudio-app.com/api/v2/ocr/jobs"
MODEL = "PaddleOCR-VL-1.6"


class OCRError(Exception):
    pass


def _get_token():
    token = os.getenv("OCR_API_TOKEN")
    if not token:
        raise OCRError("OCR_API_TOKEN is not configured in .env")
    return token


def _submit_job(file_bytes, filename):
    token = _get_token()
    headers = {"Authorization": f"bearer {token}"}
    optional_payload = {
        "useDocOrientationClassify": False,
        "useDocUnwarping": False,
        "useChartRecognition": False,
    }
    data = {
        "model": MODEL,
        "optionalPayload": json.dumps(optional_payload),
    }
    files = {"file": (filename, file_bytes)}
    resp = requests.post(JOB_URL, headers=headers, data=data, files=files)
    if resp.status_code != 200:
        raise OCRError(f"Failed to submit OCR job ({resp.status_code}): {resp.text}")
    return resp.json()["data"]["jobId"]


def _poll_job(job_id, timeout=180, interval=3):
    token = _get_token()
    headers = {"Authorization": f"bearer {token}"}
    start = time.time()
    while True:
        resp = requests.get(f"{JOB_URL}/{job_id}", headers=headers)
        if resp.status_code != 200:
            raise OCRError(f"Failed to poll OCR job ({resp.status_code}): {resp.text}")
        payload = resp.json()["data"]
        state = payload["state"]
        if state == "done":
            return payload["resultUrl"]["jsonUrl"]
        if state == "failed":
            raise OCRError(f"OCR job failed: {payload.get('errorMsg', 'unknown error')}")
        if time.time() - start > timeout:
            raise OCRError("OCR job timed out. Please try again.")
        time.sleep(interval)


def _extract_markdown(jsonl_url):
    resp = requests.get(jsonl_url)
    resp.raise_for_status()
    lines = resp.text.strip().split("\n")
    markdown_parts = []
    for line in lines:
        line = line.strip()
        if not line:
            continue
        result = json.loads(line)["result"]
        for res in result["layoutParsingResults"]:
            markdown_parts.append(res["markdown"]["text"])
    return "\n".join(markdown_parts)


def _clean_voucher_no(value):
    return re.sub(r"[^0-9]", "", str(value))


def _normalize_header(h):
    return re.sub(r"[^A-Z]", " ", str(h or "").upper()).strip()


def parse_markdown_table(markdown_text):
    """Parses the OCR markdown (HTML table) into rows with voucher, hotel, rate.
    Room Type is intentionally not read here; it's fetched from the CRM page.
    Rows with a non-numeric rate are skipped.
    """
    soup = BeautifulSoup(markdown_text, "html.parser")
    table = soup.find("table")
    if not table:
        raise OCRError("No table found in the scanned document.")

    trs = table.find_all("tr")
    if not trs:
        raise OCRError("No rows found in the scanned document's table.")

    header_cells = trs[0].find_all(["td", "th"])
    headers = [_normalize_header(re.sub(r"\s+", " ", c.get_text(strip=True))) for c in header_cells]

    def find_col(name):
        for idx, h in enumerate(headers):
            if h == name or h.startswith(name):
                return idx
        return None

    v_idx = find_col("VOUCHER")
    h_idx = find_col("HOTEL")
    r_idx = find_col("ROOM RATE")
    if r_idx is None:
        r_idx = find_col("RATE")

    # Partial match for voucher — OCR may produce "VOUCHER #", "VCHR", "VC NO", etc.
    if v_idx is None:
        for idx, h in enumerate(headers):
            if idx in (h_idx, r_idx):
                continue
            if "VOUCH" in h or "VCHR" in h or "VC" in h or "NO" in h or "#" in h:
                v_idx = idx
                break

    missing = []
    if v_idx is None:
        missing.append("VOUCHER")
    if h_idx is None:
        missing.append("HOTEL")
    if r_idx is None:
        missing.append("ROOM RATE")
    if missing:
        raise OCRError(
            f"Missing column(s) in scanned table: {', '.join(missing)}. "
            f"Found headers: {', '.join(str(h) for h in headers if h)}"
        )

    rows = []
    skipped = 0
    for tr in trs[1:]:
        cells = tr.find_all(["td", "th"])
        if len(cells) <= max(v_idx, h_idx, r_idx):
            skipped += 1
            continue

        voucher_raw = re.sub(r"\s+", " ", cells[v_idx].get_text(strip=True))
        hotel = re.sub(r"\s+", " ", cells[h_idx].get_text(strip=True))
        rate_raw = re.sub(r"\s+", " ", cells[r_idx].get_text(strip=True))

        if not voucher_raw or not hotel:
            skipped += 1
            continue

        try:
            rate_val = float(rate_raw)
        except (TypeError, ValueError):
            skipped += 1
            continue

        rows.append({
            "voucher": _clean_voucher_no(voucher_raw),
            "hotel": hotel,
            "rate": rate_val,
        })

    return rows, skipped


def image_to_rows(file_bytes, filename):
    job_id = _submit_job(file_bytes, filename)
    jsonl_url = _poll_job(job_id)
    markdown_text = _extract_markdown(jsonl_url)
    rows, skipped = parse_markdown_table(markdown_text)
    if not rows:
        raise OCRError("No valid data rows could be extracted from the image.")
    return rows, skipped
