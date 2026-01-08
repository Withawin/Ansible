import os
import sys
import json
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# =============================
# CONFIG
# =============================
SERVICE_ACCOUNT_FILE = "xxxx.json"
SPREADSHEET_NAME = "Fortigate_Inventory"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FOLDER = os.path.join(BASE_DIR, "json")
LOG_FILE = os.path.join(BASE_DIR, "upload.log")


# =============================
# GOOGLE SHEET
# =============================
def get_gsheet_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=scopes
    )
    return gspread.authorize(creds)


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
# EXTRACT JSON
# (ใช้ logic เดิมของคุณ)
# =============================
def extract_config_values(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        content = json.load(f)

    status = content.get("results", {}).get("status", {})

    SNAPSHOT_DATE = datetime.now().strftime("%Y-%m-%d")
    serial = status.get("serial", "-")

    # ----- MGMT IP -----
    mgmt_ip = "-"
    iface_mgmt = (status.get("interfacesmgmt") or [{}])[0]
    ipv4_list = iface_mgmt.get("ipv4_addresses") or []
    if ipv4_list:
        mgmt_ip = ipv4_list[0]

    # ----- NTP -----
    ntp_status = "-"
    ntp = (status.get("ntp") or [{}])[0]
    if isinstance(ntp, dict):
        ntp_status = ntp.get("ntpreachable", "-")

    fortiguard = content.get("results", {}).get("fortiguardstat", {}) or {}
    fmg = content.get("results", {}).get("fmgstat", {}) or {}
    faz = content.get("results", {}).get("fazstat", {}) or {}

    device = {
        "SnapshotDate": SNAPSHOT_DATE,
        "Hostname": status.get("hostname", "-"),
        "Serial": serial,
        "Vendor": status.get("Vendor", "Fortinet"),
        "Model": status.get("model", "-"),
        "Version": status.get("version", "-"),
        "IP Address": mgmt_ip,
        "CPU": status.get("cpu", "-"),
        "Memory": status.get("memory", "-"),
        "Uptime": status.get("uptime", "-"),
        "Session": status.get("session", "-"),
        "NTP": ntp_status,
        "Routes Total": content.get("results", {}).get("routes", {}).get("total", "-"),
        "Routes Totalv4": content.get("results", {}).get("routes", {}).get("totalv4", "-"),

        "FortiGuard Connection": fortiguard.get("fortiguardconnection", "-"),
        "FortiGuard Server": fortiguard.get("fortiguardserver", "-"),
        "FortiGuard Last Update": fortiguard.get("fortiguardlast", "-"),
        "FortiGuard Next Update": fortiguard.get("fortiguardnextupdate", "-"),

        "FMG Connection": fmg.get("fmgconnection", "-"),
        "FMG Server": fmg.get("fmgserver", "-"),
        "FMG Registration": fmg.get("fmgregistration", "-"),

        "FAZ Connection": faz.get("fazconnection", "-"),
        "FAZ IP": faz.get("fazip", "-"),
        "FAZ Registration": faz.get("fazregistration", "-"),
    }

    # ----- INTERFACES -----
    interfaces = []
    for iface in content.get("results", {}).get("interfaces", []):
        if iface.get("type") not in ["physical", "tunnel", "aggregate", "vap-switch"]:
            continue
        if iface.get("name") in ["naf.root", "l2t.root", "ssl.root", "modem", "fortilink"]:
            continue
        if iface.get("link", "").upper() != "UP":
            continue

        ip = "0.0.0.0/0"
        ipv4_list = iface.get("ipv4_addresses") or []
        if ipv4_list:
            v = ipv4_list[0]
            if isinstance(v, dict):
                ip = f"{v.get('ip')}/{v.get('cidr_netmask')}"
            elif isinstance(v, str) and v != "No IPv4 addresses available":
                ip = v

        interfaces.append({
            "SnapshotDate": SNAPSHOT_DATE,
            "Hostname": device["Hostname"],
            "Serial": serial,
            "Interface": iface.get("name"),
            "IP": ip,
            "Link": "UP"
        })

    # ----- OSPF -----
    ospf = [{
        "SnapshotDate": SNAPSHOT_DATE,
        "Hostname": device["Hostname"],
        "Serial": serial,
        "Neighbor IP": n.get("neighbor_ip"),
        "Router ID": n.get("router_id"),
        "Priority": n.get("priority")
    } for n in content.get("ospf_neighbors", [])]

    # ----- BGP -----
    bgp = [{
        "SnapshotDate": SNAPSHOT_DATE,
        "Hostname": device["Hostname"],
        "Serial": serial,
        "Neighbor IP": n.get("neighbor_ip"),
        "Local IP": n.get("local_ip"),
        "State": n.get("state")
    } for n in content.get("bgp_neighbors", [])]

    return {"device": device, "interfaces": interfaces, "ospf": ospf, "bgp": bgp}


# =============================
# MAIN
# =============================
if len(sys.argv) < 2:
    print("Usage: python process.py <hostname>")
    sys.exit(1)

hostname = sys.argv[1]
file_path = os.path.join(JSON_FOLDER, f"{hostname}_PM.json")

if not os.path.exists(file_path):
    write_log(f"Upload NOT completed for {hostname} | File not found")
    sys.exit(1)

try:
    client = get_gsheet_client()
    sh = client.open(SPREADSHEET_NAME)

    write_log(f"{hostname} | start upload")

    result = extract_config_values(file_path)

    upload_devices(sh, [result["device"]])
    write_log(f"{hostname} | devices uploaded")

    upload_interfaces(sh, result["interfaces"])
    write_log(f"{hostname} | interfaces uploaded")

    upload_ospf(sh, result["ospf"])
    write_log(f"{hostname} | ospf uploaded")

    upload_bgp(sh, result["bgp"])
    write_log(f"{hostname} | bgp uploaded")

    write_log(f"Upload completed for {hostname}")

except Exception as e:
    write_log(f"Upload NOT completed for {hostname} | {e}")
    raise
