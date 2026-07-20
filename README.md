## kagscrub

One-touch workflow to:
- Scan Outlook mailbox/folder for today's emails containing "KUNES"
- Extract VINs from subject/body
- Lookup those VINs in `raw.csv`
- Write a CSV with selected fields (same as `csv_scrubber.py`)

### Prerequisites
- Windows with Outlook installed and profile configured
- Python 3.9+ on PATH
- Install dependency:
  ```
  py -m pip install -r requirements.txt
  ```

### Quick Run (defaults configured)
```
cd "C:\Users\hz461479\OneDrive - The Hertz Corporation\Desktop\dev\kagscrub"
py outlook_kunes_lookup.py
```

Defaults:
- Mailbox: `Hertzlogistics`
- Folder: `CarMax GP`
- Input: `raw.csv` in the same directory
- Outputs:
  - `vin_list_from_outlook.txt`
  - `carmax_kunes_lookup.csv`

### Explicit Run
```
py outlook_kunes_lookup.py --raw raw.csv --mailbox "Hertzlogistics" --folder "CarMax GP" --output carmax_kunes_lookup.csv
```

### Discover Mailbox/Folder
List mailboxes:
```
py outlook_kunes_lookup.py --list-mailboxes
```

Find folder path:
```
py outlook_kunes_lookup.py --mailbox "Hertzlogistics" --discover-folder --folder-name "CarMax GP"
```

### Notes
- VIN column is assumed at index 23 in `raw.csv` (column X). Adjust `TARGET_COLUMNS` if needed.
- CSVs are ignored by git via `.gitignore` to avoid committing bulky or sensitive data. Add specific CSVs intentionally if desired.


