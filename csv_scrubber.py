#!/usr/bin/env python3
"""
CSV Scrubber Script

This script processes a raw CSV file and extracts only specific columns
for rows where the Dealership column contains "KUNES".

Output columns (matches Model-Output2.xlsx "Drop Sheet"):
  (index), vin, purchase_price, hertz_delivery_flag, hertz_name,
  hertz_address, hertz_city, hertz_state, hertz_zip,
  kmx_loc_num, kmx_address, kmx_city, kmx_state, kmx_zip

Source column mappings from raw.csv:
  vin            <- Column X  (index 23)
  purchase_price <- (not available in raw.csv; left blank)
  hertz_delivery_flag <- always 1
  hertz_name     <- Column BZ (index 77, Location Name)
  hertz_address  <- Column CE (index 82, RHP Address)
  hertz_city     <- Column CF (index 83, RHP City)
  hertz_state    <- Column CG (index 84, RHP State)
  hertz_zip      <- Column CH (index 85, RHP Postal Code)
  kmx_loc_num    <- Column J  (index  9, Dealership)
  kmx_address    <- Column N  (index 13, Address)
  kmx_city       <- Column O  (index 14, City)
  kmx_state      <- Column P  (index 15, State)
  kmx_zip        <- Column Q  (index 16, Postal Code)
"""

import csv
import sys
from pathlib import Path


def scrub_csv(input_file, output_file):
    """
    Process the CSV file and extract specific columns for KUNES dealerships.
    
    Args:
        input_file (str): Path to the input CSV file
        output_file (str): Path to the output CSV file
    """
    
    # Output column names matching Model-Output2.xlsx "Drop Sheet"
    column_names = [
        "",                    # row index (0-based)
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
    
    try:
        with open(input_file, 'r', newline='', encoding='utf-8') as infile:
            # Use csv.reader to properly handle quoted fields and commas
            reader = csv.reader(infile)
            
            with open(output_file, 'w', newline='', encoding='utf-8') as outfile:
                writer = csv.writer(outfile)
                
                # Write the header row
                writer.writerow(column_names)
                
                # Process each row
                rows_processed = 0
                rows_with_kunes = 0
                
                for row_num, row in enumerate(reader, 1):
                    rows_processed += 1
                    
                    # Skip header row for processing, but count it
                    if row_num == 1:
                        continue
                    
                    # Check if we have enough columns (highest index we read is 85 = CH)
                    if len(row) <= 85:
                        print(f"Warning: Row {row_num} has insufficient columns ({len(row)}). Skipping.")
                        continue
                    
                    # Check if Dealership column (index 9) contains "KUNES"
                    dealership = row[9].strip().upper()
                    if "KUNES" in dealership:
                        rows_with_kunes += 1
                        
                        # Build output row in target column order
                        def _col(i):
                            return row[i] if i < len(row) else ''

                        out_row = [
                            rows_with_kunes - 1,  # 0-based index
                            _col(23),             # vin (X)
                            '',                   # purchase_price (not in raw.csv)
                            1,                    # hertz_delivery_flag (always 1)
                            _col(77),             # hertz_name (BZ = Location Name)
                            _col(82),             # hertz_address (CE = RHP Address)
                            _col(83),             # hertz_city (CF = RHP City)
                            _col(84),             # hertz_state (CG = RHP State)
                            _col(85),             # hertz_zip (CH = RHP Postal Code)
                            _col(9),              # kmx_loc_num (J = Dealership)
                            _col(13),             # kmx_address (N = Address)
                            _col(14),             # kmx_city (O = City)
                            _col(15),             # kmx_state (P = State)
                            _col(16),             # kmx_zip (Q = Postal Code)
                        ]
                        writer.writerow(out_row)
                
        print(f"Processing complete!")
        print(f"Total rows processed: {rows_processed}")
        print(f"Rows with KUNES dealerships: {rows_with_kunes}")
        print(f"Output saved to: {output_file}")
                
    except FileNotFoundError:
        print(f"Error: Input file '{input_file}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error processing CSV file: {str(e)}")
        sys.exit(1)


def main():
    """Main function to run the CSV scrubber."""
    
    # Default file names
    input_file = "raw.csv"
    output_file = "scrubbed_kunes_data.csv"
    
    # Check if input file exists
    if not Path(input_file).exists():
        print(f"Error: Input file '{input_file}' not found in current directory.")
        print("Please make sure 'raw.csv' exists in the same directory as this script.")
        sys.exit(1)
    
    print(f"Starting CSV scrubbing process...")
    print(f"Input file: {input_file}")
    print(f"Output file: {output_file}")
    print(f"Looking for dealerships containing 'KUNES'...")
    print("-" * 50)
    
    scrub_csv(input_file, output_file)


if __name__ == "__main__":
    main()
