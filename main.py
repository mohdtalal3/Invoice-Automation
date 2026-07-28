import os
import re
import time
import json
import requests
import openpyxl
from collections import defaultdict
from datetime import datetime, timezone, timedelta
from dotenv import load_dotenv
from bs4 import BeautifulSoup

load_dotenv()

USERNAME = os.getenv("USER_NAME")
PASSWORD = os.getenv("PASSWORD")

BASE = "https://crm.basmaemaargroup.com"
EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "INVOICES RATES DEMO.xlsx")

# Sharing = same price, Double /2, Triple /3, Quad /4, Quint /5
ROOM_TYPE_DIVISOR = {
    "SHARING": 1,
    "DOUBLE": 2,
    "TRIPLE": 3,
    "QUAD": 4,
    "QUINT": 5,
}

COMMON_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "sec-ch-ua-platform": '"macOS"',
    "sec-ch-ua": '"Not;A=Brand";v="8", "Chromium";v="150", "Google Chrome";v="150"',
    "sec-ch-ua-mobile": "?0",
    "accept-language": "en-US,en;q=0.9,fr;q=0.8,af;q=0.7,ar;q=0.6,be;q=0.5,de;q=0.4",
    "priority": "u=1, i",
}

# Global mapping so progress/results are visible after the run
RESULTS = {}


def clean_voucher_no(value):
    return re.sub(r"[^0-9]", "", str(value))


def normalize(text):
    return re.sub(r"\s+", " ", str(text).strip().upper())


def compute_price(room_type, rate):
    divisor = ROOM_TYPE_DIVISOR.get(normalize(room_type), 1)
    return round(float(rate) / divisor, 2)


def login(session):
    now = datetime.now()
    offset = now.utcoffset()
    if offset is None:
        offset = timedelta(hours=0)
    offset_hours = int(offset.total_seconds() // 3600)
    offset_sign = "+" if offset_hours >= 0 else "-"
    offset_str = f"GMT{offset_sign}{abs(offset_hours):02d}00"
    tz_name = time.tzname[0] if time.daylight == 0 else time.tzname[1]
    dt_str = now.strftime(f"%a %b %d %Y %H:%M:%S {offset_str} ({tz_name})")
    dt_local_date = now.strftime("%d/%m/%Y")

    headers = COMMON_HEADERS.copy()
    headers.update({
        "Accept": "text/html, */*; q=0.01",
        "x-requested-with": "XMLHttpRequest",
        "content-type": "text/html; charset=utf-8",
        "sec-fetch-site": "same-origin",
        "sec-fetch-mode": "cors",
        "sec-fetch-dest": "empty",
        "referer": f"{BASE}/CRM/Login.aspx",
    })

    session.get(f"{BASE}/CRM/Login.aspx", headers=headers)

    params = {
        "Arabic": "false",
        "username": USERNAME,
        "password": PASSWORD,
        "rememberme": "0",
        "dt": dt_str,
        "TZ": "5",
        "dtLocalDate": dt_local_date,
    }
    response = session.get(f"{BASE}/CRM/LoginVerify.aspx", params=params, headers=headers)
    return "success" in response.text.lower()


def fetch_voucher_data(session, voucher_no):
    headers = COMMON_HEADERS.copy()
    headers.update({
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "x-requested-with": "XMLHttpRequest",
        "content-type": "text/html; charset=utf-8",
        "sec-fetch-site": "same-origin",
        "sec-fetch-mode": "cors",
        "sec-fetch-dest": "empty",
        "referer": f"{BASE}/CRM/TRV/Vouchers.aspx",
    })
    voucher_params = {
        "QueryType": "VoucherSearch",
        "VoucherNo": voucher_no,
        "FamilyHead": "null",
        "GroupName": "null",
        "Branch": "null",
        "Customer": "null",
        "Package": "null",
        "Status": "null",
        "MakHot": "null",
        "MadHot": "null",
        "VisaCmp": "null",
        "RoomType": "null",
        "ReturnFrom": "null",
        "ReturnTo": "null",
        "VoucherFrom": "null",
        "VoucherTo": "null",
        "CreatedFrom": "null",
        "CreatedTo": "null",
        "withoutBooking": "false",
        "Records": "TOP 50",
    }
    response = session.get(f"{BASE}/CRM/TRV/VouchersCB.aspx", params=voucher_params, headers=headers)
    try:
        return response.json()
    except ValueError:
        return None


def get_booking_page(session, booking_oid, referer):
    headers = COMMON_HEADERS.copy()
    headers.update({
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "upgrade-insecure-requests": "1",
        "sec-fetch-site": "same-origin",
        "sec-fetch-mode": "navigate",
        "sec-fetch-user": "?1",
        "sec-fetch-dest": "document",
        "referer": referer,
    })
    url = f"{BASE}/CRM/TRV/BookingAccounts.aspx"
    response = session.get(url, params={"BokOid": booking_oid}, headers=headers)
    return response.text


def post_booking_page(session, booking_oid, form_data):
    booking_url = f"{BASE}/CRM/TRV/BookingAccounts.aspx?BokOid={booking_oid}"
    headers = COMMON_HEADERS.copy()
    headers.update({
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Content-Type": "application/x-www-form-urlencoded",
        "cache-control": "max-age=0",
        "upgrade-insecure-requests": "1",
        "origin": BASE,
        "sec-fetch-site": "same-origin",
        "sec-fetch-mode": "navigate",
        "sec-fetch-user": "?1",
        "sec-fetch-dest": "document",
        "referer": booking_url,
    })
    response = session.post(booking_url, data=form_data, headers=headers)
    return response.text


def extract_form_fields(html):
    """Generic scrape of the ASP.NET form -> replicates what a real browser submits."""
    soup = BeautifulSoup(html, "html.parser")
    form = soup.find("form", id="form1") or soup.find("form")
    data = {}
    if not form:
        return data

    for inp in form.find_all("input"):
        name = inp.get("name")
        if not name:
            continue
        itype = (inp.get("type") or "text").lower()
        if itype in ("checkbox", "radio"):
            if inp.has_attr("checked"):
                data[name] = inp.get("value", "on")
        elif itype in ("submit", "button", "image", "reset", "file"):
            continue
        else:
            data[name] = inp.get("value", "")

    for sel in form.find_all("select"):
        name = sel.get("name")
        if not name:
            continue
        option = sel.find("option", selected=True) or sel.find("option")
        if option is not None:
            data[name] = option.get("value", option.get_text(strip=True))

    for ta in form.find_all("textarea"):
        name = ta.get("name")
        if name:
            data[name] = ta.get_text()

    return data


def get_hotel_divs(html):
    soup = BeautifulSoup(html, "html.parser")
    hotel_divs = []
    for div in soup.find_all("div", id=re.compile(r"^MainPane_MainContent_divHotel\d+$")):
        n = int(re.search(r"divHotel(\d+)", div["id"]).group(1))
        name_span = soup.find("span", id=f"MainPane_MainContent_lblHotel{n}")
        hotel_name = name_span.get_text(strip=True) if name_span else ""
        room_type_span = soup.find("span", id=f"MainPane_MainContent_lblRoomType{n}")
        room_type_raw = room_type_span.get_text(strip=True) if room_type_span else ""
        room_type = re.sub(r"\bBED\b", "", room_type_raw, flags=re.IGNORECASE).strip()
        hotel_divs.append({"index": n, "name": hotel_name, "room_type": room_type})
    hotel_divs.sort(key=lambda x: x["index"])
    return hotel_divs


def read_excel_rows(path):
    wb = openpyxl.load_workbook(path, data_only=True)
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


def process_voucher(session, voucher_no, excel_rows, log_fn=print, results=None):
    log_fn(f"\n=== Processing Voucher {voucher_no} ===")
    if results is None:
        results = RESULTS
    results[voucher_no] = {}

    voucher_data = fetch_voucher_data(session, voucher_no)
    if not voucher_data:
        log_fn(f"  No voucher data returned for {voucher_no}, skipping")
        results[voucher_no]["error"] = "voucher not found"
        return

    voucher_list = voucher_data if isinstance(voucher_data, list) else voucher_data.get("data", voucher_data)
    first_entry = voucher_list[0] if isinstance(voucher_list, list) and voucher_list else voucher_list
    if not first_entry:
        log_fn(f"  No entries found for voucher {voucher_no}, skipping")
        results[voucher_no]["error"] = "no entries"
        return

    booking_oid = first_entry.get("BookingOid") or first_entry.get("Oid")
    log_fn(f"  BookingOid: {booking_oid}")
    results[voucher_no]["booking_oid"] = booking_oid
    results[voucher_no]["hotels"] = []

    current_html = get_booking_page(session, booking_oid, referer=f"{BASE}/CRM/TRV/Vouchers.aspx")

    used_indices = set()

    for row in excel_rows:
        hotel_divs = get_hotel_divs(current_html)
        form_data = extract_form_fields(current_html)

        match = None
        target_name = normalize(row["hotel"])
        for hd in hotel_divs:
            if hd["index"] in used_indices:
                continue
            if normalize(hd["name"]) == target_name:
                match = hd
                break

        if not match:
            log_fn(f"  WARNING: hotel '{row['hotel']}' not found on booking page (or already matched) -> skipping")
            results[voucher_no]["hotels"].append({"hotel": row["hotel"], "status": "not_found"})
            continue

        used_indices.add(match["index"])
        n = match["index"]
        checkbox_name = f"ctl00$MainPane$MainContent$chkCustomSale{n}"
        price_name = f"ctl00$MainPane$MainContent$txtRoomPrice{n}"
        is_checked = form_data.get(checkbox_name) == "C"

        if is_checked:
            existing_price = form_data.get(price_name, "")
            log_fn(f"  [{n}] {match['name']}: invoice already exists (price={existing_price}) -> SKIP")
            results[voucher_no]["hotels"].append({
                "hotel": match["name"], "status": "already_invoiced", "price": existing_price
            })
            continue

        html_room_type = match.get("room_type")
        if not html_room_type:
            log_fn(f"  [{n}] {match['name']}: room type not found on page -> skipping")
            results[voucher_no]["hotels"].append({"hotel": match["name"], "status": "room_type_not_found"})
            continue

        price = compute_price(html_room_type, row["rate"])
        log_fn(f"  [{n}] {match['name']}: rate {row['rate']} / {html_room_type} (from page) -> setting custom price = {price}")

        form_data[checkbox_name] = "C"
        form_data[price_name] = str(price)
        form_data["ctl00$MainPane$MainContent$txtNote"] = "Done"
        form_data["ctl00$MainPane$MainContent$btnSave"] = " Update Booking "

        current_html = post_booking_page(session, booking_oid, form_data)
        log_fn(f"    -> Update submitted for {match['name']}")
        results[voucher_no]["hotels"].append({
            "hotel": match["name"], "status": "invoiced", "price": price
        })


def main():
    session = requests.Session()

    if not login(session):
        print("Login failed")
        return
    print("Account logged in successfully")

    excel_rows = read_excel_rows(EXCEL_PATH)

    unique_vouchers = sorted(set(r["voucher"] for r in excel_rows))
    print(f"\nFound {len(unique_vouchers)} unique voucher number(s) in Excel: {unique_vouchers}")

    grouped = defaultdict(list)
    for r in excel_rows:
        grouped[r["voucher"]].append(r)

    for voucher_no in unique_vouchers:
        process_voucher(session, voucher_no, grouped[voucher_no])

    print("\n\n=== SUMMARY ===")
    print(json.dumps(RESULTS, indent=2))

    with open("results_summary.json", "w") as f:
        json.dump(RESULTS, f, indent=2)
    print("\nSummary saved to results_summary.json")


if __name__ == "__main__":
    main()
