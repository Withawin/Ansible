import os
import sys
import json
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

SERVICE_ACCOUNT_FILE = "xxxx.json"
SPREADSHEET_NAME = "Fortigate_Inventory"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "upload.log")
JSON_FOLDER = os.path.join(BASE_DIR, "json")


# =============================
# GOOGLE SHEET (เปิดครั้งเดียว)
# =============================
def get_gsheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=scopes
    )
    client = gspread.authorize(creds)
    return client.open(SPREADSHEET_NAME)


# =============================
# LOG
# =============================
def write_log(message):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{ts}] {message}\n")


# =============================
# UPLOAD FUNCTIONS
# =============================
def upload_devices(sh, data):
    sheet = sh.worksheet("devices")

    headers = [
        "SnapshotDate", "createdAt", "Hostname", "Serial", "Vendor", "Model",
        "Version", "IP Address", "CPU", "Memory",
        "Uptime", "Session", "NTP",
        "Routes Total", "Routes Totalv4",
        "FortiGuard Connection", "FortiGuard Server",
        "FortiGuard Last Update", "FortiGuard Next Update",
        "FMG Connection", "FMG Server", "FMG Registration",
        "FAZ Connection", "FAZ IP", "FAZ Registration"
    ]

    if not sheet.get_all_values():
        sheet.append_row(headers)

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = []

    for d in data:
        rows.append([
            d.get("SnapshotDate"), now,
            d.get("Hostname"), d.get("Serial"),
            d.get("Vendor"), d.get("Model"),
            d.get("Version"), d.get("IP Address"),
            d.get("CPU"), d.get("Memory"),
            d.get("Uptime"), d.get("Session"),
            d.get("NTP"),
            d.get("Routes Total"), d.get("Routes Totalv4"),
            d.get("FortiGuard Connection"), d.get("FortiGuard Server"),
            d.get("FortiGuard Last Update"), d.get("FortiGuard Next Update"),
            d.get("FMG Connection"), d.get("FMG Server"),
            d.get("FMG Registration"),
            d.get("FAZ Connection"), d.get("FAZ IP"),
            d.get("FAZ Registration")
        ])

    sheet.append_rows(rows, value_input_option="RAW")


def upload_interfaces(sh, data):
    if not data:
        return

    sheet = sh.worksheet("interfaces")
    headers = ["SnapshotDate", "createdAt", "Hostname", "Serial", "Interface", "IP", "Link"]

    if not sheet.get_all_values():
        sheet.append_row(headers)

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = [[
        i.get("SnapshotDate"), now,
        i.get("Hostname"), i.get("Serial"),
        i.get("Interface"), i.get("IP"),
        i.get("Link")
    ] for i in data]

    sheet.append_rows(rows, value_input_option="RAW")


def upload_ospf(sh, data):
    if not data:
        return

    sheet = sh.worksheet("ospf_neighbors")
    headers = ["SnapshotDate", "createdAt", "Hostname", "Serial",
               "Neighbor IP", "Router ID", "Priority"]

    if not sheet.get_all_values():
        sheet.append_row(headers)

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = [[
        o.get("SnapshotDate"), now,
        o.get("Hostname"), o.get("Serial"),
        o.get("Neighbor IP"),
        o.get("Router ID"),
        o.get("Priority")
    ] for o in data]

    sheet.append_rows(rows, value_input_option="RAW")


def upload_bgp(sh, data):
    if not data:
        return

    sheet = sh.worksheet("bgp_neighbors")
    headers = ["SnapshotDate", "createdAt", "Hostname", "Serial",
               "Neighbor IP", "Local IP", "State"]

    if not sheet.get_all_values():
        sheet.append_row(headers)

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = [[
        b.get("SnapshotDate"), now,
        b.get("Hostname"), b.get("Serial"),
        b.get("Neighbor IP"),
        b.get("Local IP"),
        b.get("State")
    ] for b in data]

    sheet.append_rows(rows, value_input_option="RAW")


# =============================
# MAIN
# =============================
if len(sys.argv) < 2:
    write_log("No hostname provided")
    sys.exit(0)

hostname = sys.argv[1]
file_path = os.path.join(JSON_FOLDER, f"{hostname}_PM.json")

if not os.path.exists(file_path):
    write_log(f"Upload NOT completed for {hostname} | File not found")
    sys.exit(0)

try:
    write_log(f"{hostname} | start upload")

    sh = get_gsheet()   # ✅ เปิด Spreadsheet ครั้งเดียว
    result = extract_config_values(file_path)

    upload_devices(sh, [result["device"]])
    upload_interfaces(sh, result["interfaces"])
    upload_ospf(sh, result["ospf"])
    upload_bgp(sh, result["bgp"])

    write_log(f"Upload completed for {hostname}")

except Exception as e:
    write_log(f"Upload NOT completed for {hostname} | {e}")
