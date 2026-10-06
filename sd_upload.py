#!/usr/bin/env python3
"""
Super Dispatch Order Upload Integration

Combines Outlook gate-pass email scanning, raw CSV matching, order
consolidation, and Super Dispatch API upload into a single workflow.

Workflow:
  1. Scan Outlook "CarMax GP" folder for today's KUNES gate-pass emails
  2. Extract VINs and save gate-pass attachments to ./gate_passes/
  3. Match VINs against raw.csv to get pickup/delivery info
  4. Consolidate records with identical pickup+delivery into single orders
  5. Upload orders to Super Dispatch via API (attach gate passes per order)

Usage:
    python sd_upload.py                           # Full run, upload first order only
    python sd_upload.py --dry-run                 # Build orders, don't upload
    python sd_upload.py --use-cached-vins         # Skip Outlook, use vin_list_from_outlook.txt
    python sd_upload.py --upload-all              # Upload ALL orders (not just first)
    python sd_upload.py --order-prefix KAG60409   # Override auto-generated prefix
"""

import argparse
import csv
import json
import os
import re
import sys
from collections import defaultdict
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

try:
    import requests
except ImportError:
    print("Error: requests is required. Install with: pip install requests")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    openpyxl = None  # type: ignore[assignment]

# ---------------------------------------------------------------------------
# Outlook helpers – imported at runtime only when needed (Windows + pywin32)
# ---------------------------------------------------------------------------
_win32com = None

def _ensure_win32com():
    global _win32com
    if _win32com is not None:
        return
    try:
        import win32com.client as _w  # type: ignore
        _win32com = _w
    except Exception as exc:
        print(f"Error: pywin32 is required for Outlook access. Install with: pip install pywin32")
        print(f"Details: {exc}")
        sys.exit(1)


VIN_REGEX = re.compile(r"\b[A-HJ-NPR-Z0-9]{17}\b", re.IGNORECASE)

# ---- Super Dispatch API endpoints ----
SD_TOKEN_URL = "https://api.shipper.superdispatch.com/oauth/token"
SD_ORDERS_URL = "https://api.shipper.superdispatch.com/v1/public/orders"

# ---- Default order field values (from example upload) ----
STANDARD_INSTRUCTIONS = (
    "*For questions, contact Josh at 239-908-5189 joshua.blankenship@hertz.com "
    "or Matt at 239-301-7033 matthew.biocchi@hertz.com*\n"
    "*Pictures prior to pick up and delivery are required.*\n"
    "*Carriers please update Pickup ETA via the Super Dispatch App.*"
)

CUSTOMER_INFO = {
    "name": "Kunes Auto Group",
    "address": "1234 E Geneva St",
    "city": "Delavan",
    "state": "WI",
    "zip": "53115",
}

# ---- Raw CSV column indices (0-based) ----
COL_DEALERSHIP = 9      # J  – Dealership (delivery name)
COL_ADDRESS    = 13     # N  – Dealership address (delivery street)
COL_CITY       = 14     # O  – City (delivery)
COL_STATE      = 15     # P  – State (delivery)
COL_ZIP        = 16     # Q  – Postal Code (delivery)
COL_VIN        = 28     # AC – VIN
COL_YEAR       = 34     # AI – Year
COL_MAKE       = 35     # AJ – Make
COL_MODEL      = 36     # AK – Model
COL_LOC_NAME   = 83     # CF – Location Name (pickup name)
COL_RHP_ADDR   = 88     # CK – RHP Address (pickup street)
COL_RHP_CITY   = 89     # CL – RHP City (pickup)
COL_RHP_STATE  = 90     # CM – RHP State (pickup)
COL_RHP_ZIP    = 91     # CN – RHP Postal Code (pickup)
COL_PU_CONTACT = 132    # EC – Pickup Location Contact Name
COL_PU_EMAIL1  = 133    # ED – Pickup Location Email
COL_PU_EMAIL2  = 134    # EE – Pickup Location Email 2
COL_PU_EMAIL3  = 135    # EF – Pickup Location Email 3
COL_PU_PHONE   = 136    # EG – Pickup Location Phone Number

MAX_COL_INDEX  = 136

# Vehicle type classification based on model keywords
_SUV_KEYWORDS = {
    "ENCORE", "ENVISION", "ENCLAVE", "SPORTAGE", "SELTOS", "TELLURIDE",
    "SORENTO", "NIRO", "SOUL", "ROGUE", "PATHFINDER", "MURANO", "KICKS",
    "ARMADA", "TUCSON", "SANTA FE", "PALISADE", "KONA", "VENUE",
    "EQUINOX", "TRAX", "TRAILBLAZER", "TRAVERSE", "BLAZER", "TAHOE",
    "SUBURBAN", "EXPLORER", "ESCAPE", "BRONCO", "EDGE", "EXPEDITION",
    "ECOSPORT", "HIGHLANDER", "RAV4", "4RUNNER", "VENZA", "SEQUOIA",
    "CR-V", "HR-V", "PILOT", "PASSPORT", "WRANGLER", "CHEROKEE",
    "COMPASS", "RENEGADE", "OUTBACK", "FORESTER", "CROSSTREK", "ASCENT",
    "OUTLANDER", "ECLIPSE CROSS", "CX-5", "CX-30", "CX-50", "CX-9",
    "TIGUAN", "ATLAS", "TAOS", "ID.4", "Q5", "Q7", "X3", "X5",
    "GLC", "GLE", "RX", "NX", "UX", "MDX", "RDX", "QX",
}
_SEDAN_KEYWORDS = {
    "ALTIMA", "SENTRA", "MAXIMA", "VERSA", "CAMRY", "COROLLA",
    "ACCORD", "CIVIC", "INSIGHT", "SONATA", "ELANTRA", "ACCENT",
    "MALIBU", "IMPALA", "CRUZE", "SPARK", "FUSION", "FOCUS", "TAURUS",
    "K4", "K5", "FORTE", "OPTIMA", "STINGER", "RIO",
    "LEGACY", "IMPREZA", "WRX", "JETTA", "PASSAT", "ARTEON",
    "MAZDA3", "MAZDA6", "3 SERIES", "5 SERIES", "A4", "A6",
    "C-CLASS", "E-CLASS", "ES", "IS", "GS", "TLX", "ILX",
    "PRIUS", "LEAF", "BOLT EV", "MODEL 3", "MODEL S",
}
_PICKUP_KEYWORDS = {
    "F-150", "F-250", "F-350", "SILVERADO", "SIERRA", "RAM",
    "COLORADO", "CANYON", "TACOMA", "TUNDRA", "FRONTIER", "TITAN",
    "RANGER", "MAVERICK", "RIDGELINE", "GLADIATOR",
}
_VAN_KEYWORDS = {
    "EXPRESS", "SAVANA", "TRANSIT", "SPRINTER", "PROMASTER", "NV",
    "SIENNA", "ODYSSEY", "PACIFICA", "CARNIVAL", "CARAVAN", "METRIS",
}
_COUPE_KEYWORDS = {
    "MUSTANG", "CAMARO", "CHALLENGER", "CHARGER", "SUPRA", "BRZ",
    "86", "CORVETTE", "370Z", "400Z",
}


def classify_vehicle_type(model: str) -> str:
    """Map a vehicle model string to a Super Dispatch vehicle type enum."""
    up = model.upper().strip()
    for kw in _PICKUP_KEYWORDS:
        if kw in up:
            return "4_door_pickup" if "250" not in up and "350" not in up else "pickup"
    for kw in _VAN_KEYWORDS:
        if kw in up:
            return "van"
    for kw in _SUV_KEYWORDS:
        if kw in up:
            return "suv"
    for kw in _SEDAN_KEYWORDS:
        if kw in up:
            return "sedan"
    for kw in _COUPE_KEYWORDS:
        if kw in up:
            return "2_door_coupe"
    return "other"


# ===================================================================
#  UPLOAD LEDGER – persistent duplicate prevention
# ===================================================================

EMPTY_LEDGER: Dict[str, Any] = {"orders": {}, "vins": {}}


def load_ledger(path: Path) -> Dict[str, Any]:
    """Read upload_history.json; return empty structure if missing or corrupt."""
    if not path.exists():
        return {"orders": {}, "vins": {}}
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        if "orders" not in data or "vins" not in data:
            return {"orders": {}, "vins": {}}
        return data
    except (json.JSONDecodeError, OSError):
        return {"orders": {}, "vins": {}}


def save_ledger(path: Path, ledger: Dict[str, Any]) -> None:
    """Atomically write the ledger back to disk."""
    tmp = path.with_suffix(".tmp")
    with tmp.open("w", encoding="utf-8") as f:
        json.dump(ledger, f, indent=2)
    tmp.replace(path)


def record_upload(
    ledger: Dict[str, Any],
    order_number: str,
    guid: str,
    vins: List[str],
) -> None:
    """Add a successfully uploaded order and its VINs to the ledger."""
    ledger["orders"][order_number] = {
        "guid": guid,
        "vins": vins,
        "uploaded_at": datetime.now(timezone.utc).isoformat(),
    }
    for vin in vins:
        ledger["vins"][vin] = order_number


def filter_duplicates(
    orders: List[Dict[str, Any]],
    ledger: Dict[str, Any],
) -> Tuple[List[Dict[str, Any]], List[Tuple[str, str]]]:
    """
    Separate orders into new (uploadable) and duplicate (skipped).

    Returns:
        new_orders:  list of orders safe to upload
        skipped:     list of (order_number, reason) for duplicates
    """
    new_orders: List[Dict[str, Any]] = []
    skipped: List[Tuple[str, str]] = []

    for order in orders:
        order_num = order["payload"]["number"]
        order_vins = order["vins"]

        if order_num in ledger["orders"]:
            skipped.append((order_num, f"order number already uploaded"))
            continue

        dup_vins = [v for v in order_vins if v in ledger["vins"]]
        if dup_vins:
            prev_orders = {ledger["vins"][v] for v in dup_vins}
            skipped.append((
                order_num,
                f"VIN(s) already uploaded: {', '.join(dup_vins)} "
                f"(in {', '.join(prev_orders)})",
            ))
            continue

        new_orders.append(order)

    return new_orders, skipped


# ===================================================================
#  1.  SUPER DISPATCH AUTHENTICATION
# ===================================================================

def get_sd_access_token(client_id: str, client_secret: str) -> str:
    """Obtain an OAuth2 Bearer token from Super Dispatch."""
    resp = requests.post(
        SD_TOKEN_URL,
        data={"grant_type": "client_credentials"},
        auth=(client_id, client_secret),
    )
    if resp.status_code != 200:
        raise RuntimeError(
            f"SD auth failed ({resp.status_code}): {resp.text}"
        )
    token = resp.json().get("access_token")
    if not token:
        raise RuntimeError(f"No access_token in SD response: {resp.json()}")
    return token


# ===================================================================
#  2.  OUTLOOK – EXTRACT VINS & SAVE GATE-PASS ATTACHMENTS
# ===================================================================

def _format_outlook_dt(dt: datetime) -> str:
    return dt.strftime("%m/%d/%Y %I:%M %p")


def _open_mail_folder(namespace, mailbox_name: Optional[str], folder_path: str):
    root = None
    if mailbox_name:
        try:
            root = namespace.Folders[mailbox_name]
        except Exception:
            for store in list(namespace.Folders):
                name = getattr(store, "Name", "")
                if name and name.strip().lower() == mailbox_name.strip().lower():
                    root = store
                    break
        if root is None:
            avail = ", ".join(getattr(s, "Name", "?") for s in list(namespace.Folders))
            raise RuntimeError(f"Mailbox '{mailbox_name}' not found. Available: {avail}")
    else:
        root = namespace.GetDefaultFolder(6).Parent

    current = root
    for part in (p for p in folder_path.split("/") if p):
        try:
            current = current.Folders[part]
        except Exception:
            raise RuntimeError(f"Folder segment '{part}' not found in path '{folder_path}'.")
    return current


def collect_vins_and_gate_passes(
    mailbox: str,
    folder: str,
    gate_pass_dir: Path,
) -> Tuple[Set[str], Dict[str, List[Path]]]:
    """
    Scan Outlook for today's KUNES gate-pass emails.

    Returns:
        vins:               set of uppercase VIN strings
        vin_to_gate_passes: mapping of VIN -> list of saved file paths
    """
    _ensure_win32com()
    app = _win32com.Dispatch("Outlook.Application")
    namespace = app.GetNamespace("MAPI")

    folder_obj = _open_mail_folder(namespace, mailbox, folder)
    items = folder_obj.Items

    today_start = datetime.combine(date.today(), time.min)
    tomorrow_start = today_start + timedelta(days=1)
    filt = (
        f"[ReceivedTime] >= '{_format_outlook_dt(today_start)}' AND "
        f"[ReceivedTime] < '{_format_outlook_dt(tomorrow_start)}'"
    )
    items.Sort("[ReceivedTime]", True)
    restricted = items.Restrict(filt)

    gate_pass_dir.mkdir(parents=True, exist_ok=True)

    vins: Set[str] = set()
    vin_to_gp: Dict[str, List[Path]] = defaultdict(list)
    email_count = 0

    for item in restricted:
        try:
            subject = item.Subject or ""
            body = getattr(item, "Body", "") or ""
            if "KUNES" not in subject.upper() and "KUNES" not in body.upper():
                continue

            email_count += 1
            email_vins: Set[str] = set()
            for m in VIN_REGEX.findall(subject):
                email_vins.add(m.upper())
            for m in VIN_REGEX.findall(body):
                email_vins.add(m.upper())
            vins.update(email_vins)

            saved_files: List[Path] = []
            if item.Attachments.Count > 0:
                for i in range(1, item.Attachments.Count + 1):
                    att = item.Attachments.Item(i)
                    safe_name = re.sub(r'[<>:"/\\|?*]', "_", att.FileName)
                    dest = gate_pass_dir / safe_name
                    if not dest.exists():
                        att.SaveAsFile(str(dest))
                    saved_files.append(dest)
            else:
                msg_name = re.sub(r'[<>:"/\\|?*]', "_", subject[:80]) + ".msg"
                dest = gate_pass_dir / msg_name
                if not dest.exists():
                    item.SaveAs(str(dest), 3)  # olMSG = 3
                saved_files.append(dest)

            for vin in email_vins:
                vin_to_gp[vin].extend(saved_files)

        except Exception:
            continue

    print(f"  Emails scanned (today, KUNES): {email_count}")
    print(f"  Unique VINs extracted: {len(vins)}")
    print(f"  Gate-pass files saved: {sum(len(v) for v in vin_to_gp.values())}")
    return vins, dict(vin_to_gp)


def load_cached_vins(path: Path) -> Set[str]:
    """Read VINs from a text file (one per line)."""
    vins: Set[str] = set()
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            v = line.strip().upper()
            if VIN_REGEX.fullmatch(v):
                vins.add(v)
    return vins


# ===================================================================
#  3.  RAW CSV MATCHING
# ===================================================================

def load_kunes_delivery_overrides(xlsx_path: Path) -> Dict[str, Dict[str, str]]:
    """
    Load the kunes delivery-contact mapping from an Excel file.

    Returns a dict keyed by the raw-file dealership name (Column A, lowered)
    whose values contain the override delivery fields from Columns B-I.
    """
    if openpyxl is None:
        print("  [WARN] openpyxl not installed – skipping kunes delivery overrides.")
        print("         Install with: pip install openpyxl")
        return {}

    if not xlsx_path.exists():
        print(f"  [WARN] Kunes override file not found: {xlsx_path}")
        return {}

    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active
    overrides: Dict[str, Dict[str, str]] = {}

    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        raw_name = row[0]  # Column A – Raw File Delivery Name
        if not raw_name:
            continue
        overrides[str(raw_name).strip().lower()] = {
            "delivery_name":    str(row[1]).strip() if row[1] else "",
            "delivery_address": str(row[2]).strip() if row[2] else "",
            "delivery_city":    str(row[3]).strip() if row[3] else "",
            "delivery_state":   str(row[4]).strip() if row[4] else "",
            "delivery_zip":     str(int(row[5])).strip() if row[5] else "",
            "delivery_contact": str(row[6]).strip() if row[6] else "",
            "delivery_phone":   str(row[7]).strip() if row[7] else "",
            "delivery_email":   str(row[8]).strip() if row[8] else "",
        }

    wb.close()
    print(f"  Loaded {len(overrides)} kunes delivery override(s) from {xlsx_path.name}")
    return overrides


def _col(row: list, idx: int) -> str:
    return row[idx].strip() if idx < len(row) else ""


def load_matched_records(
    raw_csv: Path,
    vins: Set[str],
    delivery_overrides: Optional[Dict[str, Dict[str, str]]] = None,
) -> List[Dict[str, str]]:
    """
    Read raw.csv and return dicts for rows whose VIN (col AC / index 28) is in *vins*.

    When *delivery_overrides* is provided, any row whose Dealership (col J)
    matches a key in the lookup will have its delivery fields replaced with
    the override values (name, address, city, state, zip, contact, phone, email).
    """
    if delivery_overrides is None:
        delivery_overrides = {}

    records: List[Dict[str, str]] = []
    override_count = 0

    with raw_csv.open("r", newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        for row_idx, row in enumerate(reader):
            if row_idx == 0:
                continue  # skip header
            if len(row) <= MAX_COL_INDEX:
                continue
            vin = _col(row, COL_VIN).upper()
            if vin and vin in vins:
                raw_pickup_name = _col(row, COL_LOC_NAME)
                name_upper = raw_pickup_name.upper()
                if "MANHEIM" not in name_upper and "ADESA" not in name_upper:
                    raw_pickup_name = f"Hertz - {raw_pickup_name}"

                emails = [e for e in (
                    _col(row, COL_PU_EMAIL1),
                    _col(row, COL_PU_EMAIL2),
                    _col(row, COL_PU_EMAIL3),
                ) if e]

                dealership = _col(row, COL_DEALERSHIP)
                override = delivery_overrides.get(dealership.lower())

                if override:
                    override_count += 1
                    rec = {
                        "vin":              vin,
                        "year":             _col(row, COL_YEAR),
                        "make":             _col(row, COL_MAKE),
                        "model":            _col(row, COL_MODEL),
                        "pickup_name":      raw_pickup_name,
                        "pickup_address":   _col(row, COL_RHP_ADDR),
                        "pickup_city":      _col(row, COL_RHP_CITY),
                        "pickup_state":     _col(row, COL_RHP_STATE),
                        "pickup_zip":       _col(row, COL_RHP_ZIP),
                        "pickup_contact":   _col(row, COL_PU_CONTACT),
                        "pickup_email":     emails[0] if emails else "",
                        "pickup_phone":     _col(row, COL_PU_PHONE),
                        "delivery_name":    override["delivery_name"],
                        "delivery_address": override["delivery_address"],
                        "delivery_city":    override["delivery_city"],
                        "delivery_state":   override["delivery_state"],
                        "delivery_zip":     override["delivery_zip"],
                        "delivery_contact": override["delivery_contact"],
                        "delivery_phone":   override["delivery_phone"],
                        "delivery_email":   override["delivery_email"],
                    }
                else:
                    rec = {
                        "vin":              vin,
                        "year":             _col(row, COL_YEAR),
                        "make":             _col(row, COL_MAKE),
                        "model":            _col(row, COL_MODEL),
                        "pickup_name":      raw_pickup_name,
                        "pickup_address":   _col(row, COL_RHP_ADDR),
                        "pickup_city":      _col(row, COL_RHP_CITY),
                        "pickup_state":     _col(row, COL_RHP_STATE),
                        "pickup_zip":       _col(row, COL_RHP_ZIP),
                        "pickup_contact":   _col(row, COL_PU_CONTACT),
                        "pickup_email":     emails[0] if emails else "",
                        "pickup_phone":     _col(row, COL_PU_PHONE),
                        "delivery_name":    dealership,
                        "delivery_address": _col(row, COL_ADDRESS),
                        "delivery_city":    _col(row, COL_CITY),
                        "delivery_state":   _col(row, COL_STATE),
                        "delivery_zip":     _col(row, COL_ZIP),
                        "delivery_contact": "",
                        "delivery_phone":   "",
                        "delivery_email":   "",
                    }

                records.append(rec)

    if override_count:
        print(f"  Kunes delivery overrides applied: {override_count}")
    return records


# ===================================================================
#  4.  ORDER CONSOLIDATION
# ===================================================================

def _location_key(rec: Dict[str, str]) -> tuple:
    """Unique key for grouping records that share the same pickup+delivery."""
    return (
        rec["pickup_name"],  rec["pickup_address"],
        rec["pickup_city"],  rec["pickup_state"],  rec["pickup_zip"],
        rec["delivery_name"], rec["delivery_address"],
        rec["delivery_city"], rec["delivery_state"], rec["delivery_zip"],
    )


def consolidate_orders(
    records: List[Dict[str, str]],
    order_prefix: str,
    vin_to_gp: Dict[str, List[Path]],
    start_number: int = 1,
) -> List[Dict[str, Any]]:
    """
    Group records with matching pickup+delivery into consolidated orders.
    Returns list of order dicts ready for the SD API.
    """
    groups: Dict[tuple, List[Dict[str, str]]] = defaultdict(list)
    for rec in records:
        groups[_location_key(rec)].append(rec)

    orders: List[Dict[str, Any]] = []
    order_num = start_number

    for key, group in groups.items():
        first = group[0]
        order_id = f"{order_prefix}{order_num:03d}"

        vehicles = []
        gate_pass_files: List[Path] = []
        for rec in group:
            v: Dict[str, Any] = {
                "vin": rec["vin"],
                "is_inoperable": False,
                "type": classify_vehicle_type(rec["model"]),
            }
            vehicles.append(v)
            gate_pass_files.extend(vin_to_gp.get(rec["vin"], []))

        today = date.today()
        is_friday = today.weekday() == 4

        if is_friday:
            monday = today + timedelta(days=3)
            pickup_start = datetime(monday.year, monday.month, monday.day, 9, tzinfo=timezone.utc).strftime("%Y-%m-%dT09:00:00.000+0000")
            pickup_end = datetime(monday.year, monday.month, monday.day, 18, tzinfo=timezone.utc).strftime("%Y-%m-%dT18:00:00.000+0000")
            pickup_end = (datetime(monday.year, monday.month, monday.day, 18, tzinfo=timezone.utc) + timedelta(days=3)).strftime("%Y-%m-%dT18:00:00.000+0000")
            delivery_start = (datetime(monday.year, monday.month, monday.day, 9, tzinfo=timezone.utc) + timedelta(days=1)).strftime("%Y-%m-%dT09:00:00.000+0000")
            delivery_end = (datetime(monday.year, monday.month, monday.day, 18, tzinfo=timezone.utc) + timedelta(days=5)).strftime("%Y-%m-%dT18:00:00.000+0000")
        else:
            now = datetime.now(timezone.utc)
            pickup_start = now.strftime("%Y-%m-%dT09:00:00.000+0000")
            pickup_end = (now + timedelta(days=2)).strftime("%Y-%m-%dT18:00:00.000+0000")
            delivery_start = (now + timedelta(days=1)).strftime("%Y-%m-%dT09:00:00.000+0000")
            delivery_end = (now + timedelta(days=5)).strftime("%Y-%m-%dT18:00:00.000+0000")

        order_payload: Dict[str, Any] = {
            "number": order_id,
            "transport_type": "OPEN",
            "inspection_type": "advanced",
            "instructions": STANDARD_INSTRUCTIONS,
            "dispatcher_name": "Matt Biocchi",
            "sales_representative": "Josh Blankenship",
            "payment": {"method": "ach", "terms": "2_days"},
            "tags": ["CSRM"],
            "customer": {
                "name": CUSTOMER_INFO["name"],
                "address": CUSTOMER_INFO["address"],
                "city": CUSTOMER_INFO["city"],
                "state": CUSTOMER_INFO["state"],
                "zip": CUSTOMER_INFO["zip"],
            },
            "pickup": {
                "date_type": "estimated",
                "scheduled_at": pickup_start,
                "scheduled_ends_at": pickup_end,
                "venue": {
                    "name":    first["pickup_name"],
                    "address": first["pickup_address"],
                    "city":    first["pickup_city"],
                    "state":   first["pickup_state"],
                    "zip":     first["pickup_zip"],
                    **{k: v for k, v in {
                        "contact_name":  first.get("pickup_contact", ""),
                        "contact_email": first.get("pickup_email", ""),
                        "contact_phone": first.get("pickup_phone", ""),
                    }.items() if v},
                },
            },
            "delivery": {
                "date_type": "estimated",
                "scheduled_at": delivery_start,
                "scheduled_ends_at": delivery_end,
                "venue": {
                    "name":    first["delivery_name"],
                    "address": first["delivery_address"],
                    "city":    first["delivery_city"],
                    "state":   first["delivery_state"],
                    "zip":     first["delivery_zip"],
                    **{k: v for k, v in {
                        "contact_name":  first.get("delivery_contact", ""),
                        "contact_phone": first.get("delivery_phone", ""),
                        "contact_email": first.get("delivery_email", ""),
                    }.items() if v},
                },
            },
            "vehicles": vehicles,
        }

        orders.append({
            "payload": order_payload,
            "gate_pass_files": list(set(gate_pass_files)),
            "vins": [r["vin"] for r in group],
        })
        order_num += 1

    return orders


# ===================================================================
#  5.  SUPER DISPATCH API – CREATE ORDER & ATTACH FILES
# ===================================================================

def create_sd_order(token: str, payload: dict) -> requests.Response:
    """POST a new order to Super Dispatch."""
    return requests.post(
        SD_ORDERS_URL,
        json=payload,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json;charset=UTF-8",
        },
    )


def attach_gate_pass(token: str, order_guid: str, file_path: Path) -> Optional[requests.Response]:
    """
    Attempt to upload a gate-pass file as an order attachment.
    Returns the response, or None if the file doesn't exist.
    """
    if not file_path.exists():
        print(f"  [WARN] Gate-pass file not found: {file_path}")
        return None

    url = f"{SD_ORDERS_URL}/{order_guid}/attachments"
    with file_path.open("rb") as fh:
        resp = requests.post(
            url,
            files={"file": (file_path.name, fh)},
            headers={"Authorization": f"Bearer {token}"},
        )
    return resp


# ===================================================================
#  6.  REPORTING
# ===================================================================

def print_order_summary(orders: List[Dict[str, Any]]) -> None:
    print(f"\n{'='*60}")
    print(f"  CONSOLIDATED ORDERS: {len(orders)}")
    print(f"{'='*60}")
    for i, order in enumerate(orders, 1):
        p = order["payload"]
        vins = order["vins"]
        gp = order["gate_pass_files"]
        print(f"\n  Order {i}: {p['number']}")
        print(f"    Pickup:   {p['pickup']['venue']['name']}")
        print(f"              {p['pickup']['venue']['address']}, "
              f"{p['pickup']['venue']['city']}, "
              f"{p['pickup']['venue']['state']} "
              f"{p['pickup']['venue']['zip']}")
        print(f"    Delivery: {p['delivery']['venue']['name']}")
        print(f"              {p['delivery']['venue']['address']}, "
              f"{p['delivery']['venue']['city']}, "
              f"{p['delivery']['venue']['state']} "
              f"{p['delivery']['venue']['zip']}")
        print(f"    Vehicles ({len(vins)}):")
        for vin in vins:
            print(f"      - {vin}")
        if gp:
            print(f"    Gate passes ({len(gp)}):")
            for f in gp:
                print(f"      - {f.name}")
        else:
            print(f"    Gate passes: none saved")


# ===================================================================
#  7.  MAIN
# ===================================================================

def _next_business_day() -> date:
    """Return the next business day: Monday if today is Friday, otherwise tomorrow."""
    today = date.today()
    if today.weekday() == 4:  # Friday
        return today + timedelta(days=3)
    return today + timedelta(days=1)


def generate_order_prefix() -> str:
    """KAG60 + next business day's date as {month}{day:02d}."""
    nbd = _next_business_day()
    return f"KAG60{nbd.month}{nbd.day:02d}"


def next_order_number(ledger: Dict[str, Any], prefix: str) -> int:
    """Find the highest order number in the ledger for *prefix* and return the next one."""
    max_num = 0
    for order_id in ledger.get("orders", {}):
        if order_id.startswith(prefix):
            suffix = order_id[len(prefix):]
            try:
                num = int(suffix)
                if num > max_num:
                    max_num = num
            except ValueError:
                continue
    return max_num + 1


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Build and upload Super Dispatch orders from Outlook gate passes + raw.csv"
    )
    p.add_argument("--client-id", default="AGYMT8uasBNJr0WLMFi9",
                    help="Super Dispatch OAuth client_id")
    p.add_argument("--client-secret", default="H7mSg65rGaqAdZ8mA18eOpIGN5snLEwxVQgJyoYB",
                    help="Super Dispatch OAuth client_secret")
    p.add_argument("--raw", default="raw.csv",
                    help="Path to raw CSV (default: raw.csv)")
    p.add_argument("--mailbox", default="Hertzlogistics",
                    help="Outlook mailbox name")
    p.add_argument("--folder", default="CarMax GP",
                    help="Outlook folder path within mailbox")
    p.add_argument("--gate-pass-dir", default="gate_passes",
                    help="Directory to save gate-pass files (default: gate_passes/)")
    p.add_argument("--use-cached-vins", action="store_true",
                    help="Skip Outlook; use vin_list_from_outlook.txt instead")
    p.add_argument("--vin-file", default="vin_list_from_outlook.txt",
                    help="Cached VIN list file (used with --use-cached-vins)")
    p.add_argument("--kunes-file", default="kunes upload contacts for api.xlsx",
                    help="Kunes delivery-contact override Excel file (default: kunes upload contacts for api.xlsx)")
    p.add_argument("--order-prefix",
                    help="Override order prefix (default: auto-generated KAG60{date})")
    p.add_argument("--dry-run", action="store_true",
                    help="Build orders and print summary without uploading")
    p.add_argument("--upload-all", action="store_true",
                    help="Upload ALL orders instead of just the first one")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    script_dir = Path(__file__).resolve().parent
    raw_csv = Path(args.raw) if Path(args.raw).is_absolute() else script_dir / args.raw
    gate_pass_dir = script_dir / args.gate_pass_dir
    order_prefix = args.order_prefix or generate_order_prefix()

    print(f"Order prefix: {order_prefix}")
    print(f"Raw CSV:      {raw_csv}")

    # -- Step 1: Extract VINs (and optionally gate passes) --
    vin_to_gp: Dict[str, List[Path]] = {}

    if args.use_cached_vins:
        vin_file = Path(args.vin_file) if Path(args.vin_file).is_absolute() else script_dir / args.vin_file
        print(f"\n[1] Loading cached VINs from {vin_file} ...")
        vins = load_cached_vins(vin_file)
        print(f"  VINs loaded: {len(vins)}")
    else:
        print(f"\n[1] Scanning Outlook ({args.mailbox} / {args.folder}) for today's KUNES gate passes ...")
        vins, vin_to_gp = collect_vins_and_gate_passes(
            args.mailbox, args.folder, gate_pass_dir
        )

    if not vins:
        print("\nNo VINs found. Nothing to process.")
        return

    # -- Step 1b: Load kunes delivery overrides --
    kunes_path = Path(args.kunes_file) if Path(args.kunes_file).is_absolute() else script_dir / args.kunes_file
    print(f"\n[1b] Loading kunes delivery overrides from {kunes_path.name} ...")
    delivery_overrides = load_kunes_delivery_overrides(kunes_path)

    # -- Step 2: Match VINs to raw.csv --
    print(f"\n[2] Matching {len(vins)} VINs against {raw_csv.name} ...")
    records = load_matched_records(raw_csv, vins, delivery_overrides)
    print(f"  Matched records: {len(records)}")

    if not records:
        print("\nNo matching records in raw.csv. Ensure the raw file is up to date.")
        return

    # -- Step 3: Consolidate into orders --
    ledger_path = script_dir / "upload_history.json"
    ledger = load_ledger(ledger_path)
    start_num = next_order_number(ledger, order_prefix)
    print(f"\n[3] Consolidating orders (grouping by pickup + delivery) ...")
    print(f"  Starting order number: {order_prefix}{start_num:03d}")
    orders = consolidate_orders(records, order_prefix, vin_to_gp, start_number=start_num)
    print_order_summary(orders)

    # -- Step 4: Duplicate check --
    new_orders, skipped = filter_duplicates(orders, ledger)

    if skipped:
        print(f"\n{'='*60}")
        print(f"  DUPLICATES SKIPPED: {len(skipped)}")
        print(f"{'='*60}")
        for order_num, reason in skipped:
            print(f"    {order_num} -- {reason}")

    print(f"\n  New orders eligible for upload: {len(new_orders)}")
    print(f"  Duplicate orders skipped:       {len(skipped)}")

    if not new_orders:
        print("\nAll orders are duplicates. Nothing to upload.")
        return

    # -- Step 5: Upload to Super Dispatch --
    if args.dry_run:
        print("\n[DRY RUN] Skipping API upload.")
        print("\nSample payload for first new order:")
        print(json.dumps(new_orders[0]["payload"], indent=2))
        return

    print(f"\n[5] Authenticating with Super Dispatch ...")
    try:
        token = get_sd_access_token(args.client_id, args.client_secret)
        print("  Authentication successful.")
    except Exception as exc:
        print(f"  Authentication FAILED: {exc}")
        return

    orders_to_upload = new_orders if args.upload_all else new_orders[:1]
    print(f"\n[6] Uploading {len(orders_to_upload)} order(s) to Super Dispatch ...\n")

    uploaded: List[Dict[str, Any]] = []
    failed: List[str] = []

    for order in orders_to_upload:
        payload = order["payload"]
        order_id = payload["number"]
        print(f"  --- {order_id} ---")
        print(f"  Payload:\n{json.dumps(payload, indent=2)}\n")

        resp = create_sd_order(token, payload)
        print(f"  HTTP {resp.status_code}")

        if resp.status_code in (200, 201):
            resp_data = resp.json()
            order_guid = (
                resp_data.get("data", {}).get("object", {}).get("guid")
                or resp_data.get("data", {}).get("guid")
                or resp_data.get("guid")
            )
            print(f"  Order created successfully!  GUID: {order_guid}")

            record_upload(ledger, order_id, order_guid or "", order["vins"])
            save_ledger(ledger_path, ledger)
            uploaded.append(order)

            if order_guid and order.get("gate_pass_files"):
                print(f"  Attaching {len(order['gate_pass_files'])} gate-pass file(s) ...")
                for gp_file in order["gate_pass_files"]:
                    att_resp = attach_gate_pass(token, order_guid, gp_file)
                    if att_resp is not None:
                        if att_resp.status_code in (200, 201):
                            print(f"    Attached: {gp_file.name}")
                        else:
                            print(f"    Attach failed ({att_resp.status_code}): {att_resp.text[:200]}")
        else:
            print(f"  Order creation FAILED: {resp.text[:500]}")
            failed.append(order_id)
        print()

    remaining = len(new_orders) - len(orders_to_upload)
    if remaining > 0:
        print(f"[INFO] {remaining} additional new order(s) were built but NOT uploaded (test mode).")
        print("       Re-run with --upload-all to upload everything.\n")

    # -- Upload summary --
    total_vins = sum(len(o["vins"]) for o in uploaded)
    total_gps = sum(len(o["gate_pass_files"]) for o in uploaded)
    ledger_total_orders = len(ledger.get("orders", {}))
    ledger_total_vins = len(ledger.get("vins", {}))

    print(f"\n{'='*60}")
    print(f"  UPLOAD SUMMARY")
    print(f"{'='*60}")
    print(f"  Orders uploaded (this run): {len(uploaded)}")
    print(f"  VINs uploaded (this run):   {total_vins}")
    print(f"  Gate passes attached:       {total_gps}")
    if failed:
        print(f"  Orders failed:              {len(failed)} ({', '.join(failed)})")
    if skipped:
        print(f"  Duplicates skipped:         {len(skipped)}")
    print(f"  ----------------------------------------")
    print(f"  Total orders (all time):    {ledger_total_orders}")
    print(f"  Total VINs (all time):      {ledger_total_vins}")
    print()
    if uploaded:
        for o in uploaded:
            p = o["payload"]
            print(f"    {p['number']}  |  {len(o['vins'])} VIN(s)  |  "
                  f"{p['pickup']['venue']['name']} -> {p['delivery']['venue']['name']}")
            for vin in o["vins"]:
                print(f"      {vin}")
    print(f"{'='*60}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrupted by user.")
        sys.exit(130)
    except Exception as exc:
        print(f"\nFatal error: {exc}")
        sys.exit(1)
