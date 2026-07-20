#!/usr/bin/env python3
"""
Outlook KUNES VIN Lookup

This script connects to Outlook on Windows, scans the mailbox folder
Inbox/CarMax GP under the mailbox hertzlogistics@hertz.com for today's emails
containing the keyword "KUNES", extracts VINs from the subject/body, and then
looks up those VINs in the provided raw CSV file (expected VIN column at index 23).

Outputs:
- A CSV containing the same selected columns as csv_scrubber.py for rows whose VIN matches any extracted VIN:
  J (Dealership), N (Address), O (City), P (State), Q (Postal Code), X (VIN),
  BZ (Location Name), CE (RHP Address), CF (RHP City), CG (RHP State), CH (RHP Postal Code)
- A text file listing all extracted VINs (deduplicated)

Requirements:
    pip install pywin32

Quick run (uses your defaults):
    python outlook_kunes_lookup.py

Explicit run:
    python outlook_kunes_lookup.py --raw raw.csv \
        --mailbox "Hertzlogistics" \
        --folder "CarMax GP" \
        --output carmax_kunes_lookup.csv

Folder discovery:
    python outlook_kunes_lookup.py --mailbox "hertzlogistics@hertz.com" --discover-folder --folder-name "CarMax GP"

List available Outlook mailboxes (top-level stores):
    python outlook_kunes_lookup.py --list-mailboxes
"""

import argparse
import csv
import re
import sys
from datetime import datetime, date, time, timedelta
from pathlib import Path
from typing import Iterable, Optional, Set, Tuple

try:
    import win32com.client  # type: ignore
except Exception as import_error:  # pragma: no cover
    print("Error: pywin32 is required to access Outlook. Install with: pip install pywin32")
    print(f"Details: {import_error}")
    sys.exit(1)


VIN_REGEX = re.compile(r"\b[A-HJ-NPR-Z0-9]{17}\b", re.IGNORECASE)

# Columns to extract (0-based indices) matching csv_scrubber.py
# J=9, N=13, O=14, P=15, Q=16, X=23, BZ=77, CE=82, CF=83, CG=84, CH=85
#
# Output column order (matches Model-Output2.xlsx "Drop Sheet"):
#   (index), vin, purchase_price, hertz_delivery_flag, hertz_name,
#   hertz_address, hertz_city, hertz_state, hertz_zip,
#   kmx_loc_num, kmx_address, kmx_city, kmx_state, kmx_zip
#
# Raw CSV source indices for each output field:
#   vin            <- X=23
#   hertz_name     <- BZ=77
#   hertz_address  <- CE=82
#   hertz_city     <- CF=83
#   hertz_state    <- CG=84
#   hertz_zip      <- CH=85
#   kmx_loc_num    <- J=9  (Dealership / Kunes dealer name)
#   kmx_address    <- N=13
#   kmx_city       <- O=14
#   kmx_state      <- P=15
#   kmx_zip        <- Q=16
TARGET_COLUMNS = [9, 13, 14, 15, 16, 23, 77, 82, 83, 84, 85]  # kept for raw-CSV extraction
COLUMN_NAMES = [
    "",                   # row index (0-based)
    "vin",
    "purchase_price",
    "hertz_delivery_flag",
    "hertz_name",
    "hertz_address",
    "hertz_city",
    "hertz_state",
    "hertz_zip",
    "kmx_loc_num",
    "kmx_address",
    "kmx_city",
    "kmx_state",
    "kmx_zip",
]


def format_outlook_datetime(dt: datetime) -> str:
    """Format datetime for Outlook Restrict filters (MM/DD/YYYY HH:MM AM/PM)."""
    return dt.strftime("%m/%d/%Y %I:%M %p")


def get_outlook_namespace():
    """Return Outlook MAPI namespace."""
    application = win32com.client.Dispatch("Outlook.Application")
    return application.GetNamespace("MAPI")


def open_mail_folder(namespace, mailbox_name: Optional[str], folder_path: str):
    """
    Open a subfolder by path within a mailbox.

    folder_path example: "Inbox/CarMax GP"
    If mailbox_name is None, tries the default store; otherwise opens the top-level
    store by the provided mailbox name (e.g., "hertzlogistics@hertz.com").
    """
    root = None
    try:
        if mailbox_name:
            # Try exact match first
            try:
                root = namespace.Folders[mailbox_name]
            except Exception:
                # Try case-insensitive match across available stores
                for store in list(namespace.Folders):
                    name = getattr(store, "Name", "")
                    if name and name.strip().lower() == mailbox_name.strip().lower():
                        root = store
                        break
            if root is None:
                available = ", ".join([getattr(s, "Name", "<unknown>") for s in list(namespace.Folders)])
                raise RuntimeError(
                    f"Outlook mailbox '{mailbox_name}' not found. Available mailboxes: {available}"
                )
        else:
            root = namespace.GetDefaultFolder(6).Parent  # 6 = olFolderInbox
    except Exception:
        available = ", ".join([getattr(s, "Name", "<unknown>") for s in list(namespace.Folders)])
        raise RuntimeError(
            f"Outlook mailbox '{mailbox_name}' not found. Available mailboxes: {available}"
        )

    current = root
    for part in [p for p in folder_path.split("/") if p]:
        try:
            current = current.Folders[part]
        except Exception:
            raise RuntimeError(
                f"Outlook folder path not found: '{folder_path}'. Missing segment: '{part}'."
            )
    return current


def _walk_folder_paths(parent, prefix: str):
    """Yield (full_path, folder_obj) for all descendants of parent."""
    for sub in list(parent.Folders):
        path = sub.Name if not prefix else f"{prefix}/{sub.Name}"
        yield path, sub
        # Recurse
        try:
            yield from _walk_folder_paths(sub, path)
        except Exception:
            continue


def discover_folder_paths(namespace, mailbox_name: Optional[str], search_name: str):
    """Return list of folder paths whose name contains search_name (case-insensitive)."""
    try:
        root = namespace.Folders[mailbox_name] if mailbox_name else namespace.GetDefaultFolder(6).Parent
    except Exception:
        raise RuntimeError(
            f"Outlook mailbox '{mailbox_name}' not found. Ensure you have access in Outlook."
        )

    matches = []
    term = search_name.strip().lower()

    # Walk all top-level folders under the mailbox root
    for path, _folder in _walk_folder_paths(root, prefix=""):
        try:
            if term in (path.split("/")[-1].lower()):
                matches.append(path)
        except Exception:
            continue

    # Sort with Inbox-first preference
    matches.sort(key=lambda p: (not p.lower().startswith("inbox/"), p.lower()))
    return matches


def collect_today_kunes_vins(items) -> Tuple[Set[str], int]:
    """
    From an Outlook Items collection, restrict to today's emails, then extract VINs
    from emails whose subject/body contains 'KUNES'. Returns (vin_set, email_count).
    """
    today_start = datetime.combine(date.today(), time.min)
    tomorrow_start = today_start + timedelta(days=1)

    filter_str = (
        f"[ReceivedTime] >= '{format_outlook_datetime(today_start)}' AND "
        f"[ReceivedTime] < '{format_outlook_datetime(tomorrow_start)}'"
    )

    try:
        items.Sort("[ReceivedTime]", True)
        restricted = items.Restrict(filter_str)
    except Exception as e:
        raise RuntimeError(f"Failed to restrict Outlook items: {e}")

    vins: Set[str] = set()
    email_counter = 0

    for item in restricted:
        try:
            subject = (item.Subject or "")
            body = (getattr(item, "Body", "") or "")
            subject_upper = subject.upper()
            body_upper = body.upper()

            if "KUNES" not in subject_upper and "KUNES" not in body_upper:
                continue

            email_counter += 1

            for match in VIN_REGEX.findall(subject):
                vins.add(match.upper())
            for match in VIN_REGEX.findall(body):
                vins.add(match.upper())
        except Exception:
            # Skip non-mail items or inaccessible items gracefully
            continue

    return vins, email_counter


def lookup_vins_in_raw_csv(raw_csv_path: Path, vins: Set[str], output_csv_path: Path) -> int:
    """
    Search raw.csv for rows where VIN column (index 23) matches any VIN in 'vins'.
    Writes columns in the Model-Output2.xlsx "Drop Sheet" format:
      (index), vin, purchase_price, hertz_delivery_flag, hertz_name,
      hertz_address, hertz_city, hertz_state, hertz_zip,
      kmx_loc_num, kmx_address, kmx_city, kmx_state, kmx_zip
    Returns number of matched rows written.
    """
    if not raw_csv_path.exists():
        raise FileNotFoundError(f"Input file '{raw_csv_path}' not found.")

    matches = 0

    with raw_csv_path.open("r", newline="", encoding="utf-8") as infile, \
         output_csv_path.open("w", newline="", encoding="utf-8") as outfile:
        reader = csv.reader(infile)
        writer = csv.writer(outfile)

        # Write standardized header
        writer.writerow(COLUMN_NAMES)

        for row_index, row in enumerate(reader, start=1):
            if row_index == 1:
                # Skip the header row in raw.csv
                continue

            if len(row) <= 23:
                # Row does not have VIN column; skip
                continue

            vin_in_row = (row[23] or "").strip().upper()
            if vin_in_row and vin_in_row in vins:
                # Build output row in target column order:
                # index, vin, purchase_price, hertz_delivery_flag,
                # hertz_name, hertz_address, hertz_city, hertz_state, hertz_zip,
                # kmx_loc_num, kmx_address, kmx_city, kmx_state, kmx_zip
                def _col(i):
                    return row[i] if i < len(row) else ""

                out_row = [
                    matches,          # 0-based row index
                    _col(23),         # vin (X)
                    "",               # purchase_price (not in raw.csv)
                    1,                # hertz_delivery_flag (always 1)
                    _col(77),         # hertz_name (BZ = Location Name)
                    _col(82),         # hertz_address (CE = RHP Address)
                    _col(83),         # hertz_city (CF = RHP City)
                    _col(84),         # hertz_state (CG = RHP State)
                    _col(85),         # hertz_zip (CH = RHP Postal Code)
                    _col(9),          # kmx_loc_num (J = Dealership)
                    _col(13),         # kmx_address (N = Address)
                    _col(14),         # kmx_city (O = City)
                    _col(15),         # kmx_state (P = State)
                    _col(16),         # kmx_zip (Q = Postal Code)
                ]
                writer.writerow(out_row)
                matches += 1

    return matches


def write_vin_list(vins: Iterable[str], output_txt_path: Path) -> None:
    with output_txt_path.open("w", encoding="utf-8") as handle:
        for vin in sorted(set(vins)):
            handle.write(f"{vin}\n")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Scan Outlook CarMax GP for KUNES, extract VINs, lookup in raw.csv.")
    parser.add_argument("--mailbox", default="Hertzlogistics", help="Mailbox name/email to open (default: Hertzlogistics)")
    parser.add_argument("--folder", default="CarMax GP", help="Folder path within mailbox, e.g., 'CarMax GP'")
    parser.add_argument("--raw", default="raw.csv", help="Path to raw CSV file (default: raw.csv)")
    parser.add_argument("--output", default="carmax_kunes_lookup.csv", help="Output CSV for matched rows (default: carmax_kunes_lookup.csv)")
    parser.add_argument("--vin-output", default="vin_list_from_outlook.txt", help="Output text file listing extracted VINs")
    parser.add_argument("--discover-folder", action="store_true", help="Discover and print matching folder paths, then exit")
    parser.add_argument("--folder-name", default="CarMax GP", help="Folder name to search for when using --discover-folder")
    parser.add_argument("--list-mailboxes", action="store_true", help="List available Outlook mailboxes and exit")
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    raw_csv_path = Path(args.raw)
    output_csv_path = Path(args.output)
    vin_list_path = Path(args.vin_output)

    print("Connecting to Outlook...")
    namespace = get_outlook_namespace()

    if args.list_mailboxes:
        print("Available Outlook mailboxes (top-level stores):")
        for store in list(namespace.Folders):
            print(f" - {getattr(store, 'Name', '<unknown>')}")
        sys.exit(0)

    if args.discover_folder:
        print(f"Discovering folders matching '{args.folder_name}' in mailbox '{args.mailbox}'...")
        paths = discover_folder_paths(namespace, args.mailbox, args.folder_name)
        if not paths:
            print("No matching folders found.")
            sys.exit(2)
        print("Matches:")
        for p in paths:
            print(f" - {p}")
        # If exactly one match, provide a ready-to-use example
        if len(paths) == 1:
            print(f"\nUse this folder path with --folder:\n{paths[0]}")
        sys.exit(0)

    print(f"Opening mailbox '{args.mailbox}' and folder '{args.folder}'...")
    folder = open_mail_folder(namespace, args.mailbox, args.folder)

    print("Collecting today's KUNES emails and extracting VINs...")
    vins, email_count = collect_today_kunes_vins(folder.Items)
    print(f"Emails scanned (today, containing 'KUNES'): {email_count}")
    print(f"Unique VINs found: {len(vins)}")

    if not vins:
        print("No VINs found today in emails containing 'KUNES'. Nothing to lookup.")
        write_vin_list([], vin_list_path)
        sys.exit(0)

    print(f"Writing VIN list to: {vin_list_path}")
    write_vin_list(vins, vin_list_path)

    print(f"Looking up VINs in raw CSV: {raw_csv_path}")
    matches = lookup_vins_in_raw_csv(raw_csv_path, vins, output_csv_path)
    print(f"Matched rows written: {matches}")
    print(f"Output CSV saved to: {output_csv_path}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("Interrupted by user.")
        sys.exit(130)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


