# parse_cb.py

import os
import re
import json
import logging
import pandas as pd
from datetime import datetime

def setup_logger(log_path):
    os.makedirs(os.path.dirname(log_path), exist_ok=True)
    logger = logging.getLogger("cb_parser")
    logger.setLevel(logging.INFO)
    fh = logging.FileHandler(log_path, encoding='utf-8')
    fh.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    fh.setFormatter(formatter)
    if logger.hasHandlers():
        logger.handlers.clear()
    logger.addHandler(fh)
    return logger

def print_last_log_lines(log_path, n=10):
    if not os.path.exists(log_path):
        print("No log file found.")
        return
    with open(log_path, encoding='utf-8') as f:
        lines = f.readlines()
        print("".join(lines[-n:]))

def write_json_output(data, output_dir, original_filename, logger):
    os.makedirs(output_dir, exist_ok=True)
    base_name = os.path.splitext(original_filename)[0]
    json_filename = base_name + ".json"
    output_path = os.path.join(output_dir, json_filename)
    try:
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        logger.info(f"Output written: {output_path}")
    except Exception as e:
        logger.error(f"Failed to write JSON output: {e}")

def extract_delivery_date_from_filename(filename):
    # Match d.m, dd.mm, d.m.yy, d.m.yyyy, dd.mm.yy, dd.mm.yyyy
    match = re.search(r'(\d{1,2})[.\-](\d{1,2})(?:[.\-](\d{2,4}))?', filename)
    if not match:
        return None

    day, month, year = match.groups()
    if not year:
        year = str(datetime.now().year)
    elif len(year) == 2:
        year = '20' + year  # Assume 21st century for 2-digit years

    try:
        date_obj = datetime.strptime(f"{int(day)}.{int(month)}.{int(year)}", "%d.%m.%Y")
        return date_obj.strftime("%Y-%m-%d")
    except ValueError:
        return None


def parse_cb(file_path, path_config):
    filename = os.path.basename(file_path)

    with open("config/column_map_cb.json", encoding='utf-8') as f:
        config = json.load(f)

    LOG_PATH = os.path.join(path_config["output_cb"], "parse_cb.log")
    OUTPUT_DIR = path_config["output_cb"]
    logger = setup_logger(LOG_PATH)
    logger.info(f"Parsing file: {filename}")

    sheet_name = config["sheet_name"]
    header_row = config["header_row"] - 1
    quantity_candidates = config["columns"]["quantity_candidates"]
    product_col_name = config["columns"]["product_name"]
    unit_price_col = config["columns"].get("unit_price")
    default_tax = config["defaults"]["tax"]

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, engine='openpyxl')
    except Exception as e:
        logger.error(f"Failed to load sheet '{sheet_name}' from {filename}: {e}")
        return

    df.dropna(how='all', inplace=True)

    qty_col = None
    for col in quantity_candidates:
        if col in df.columns and pd.to_numeric(df[col], errors='coerce').notna().sum() > 0:
            qty_col = col
            break
    if not qty_col:
        logger.error(f"No valid quantity column found in {filename}")
        return

    delivery_date = extract_delivery_date_from_filename(filename)
    if not delivery_date:
        logger.warning(f"Delivery date set to null for {filename}")

    rows = []
    for idx, row in df.iterrows():
        product_name = str(row.get(product_col_name)).strip()
        if not product_name or product_name.lower() in ['nan', 'tổng cộng']:
            continue

        try:
            qty = float(row[qty_col])
        except (ValueError, TypeError):
            continue

        unit_price = None
        if unit_price_col and unit_price_col in df.columns:
            try:
                unit_price = float(row[unit_price_col])
            except (ValueError, TypeError):
                unit_price = None

        rows.append({
            "product_name": product_name,
            "qty": qty,
            "unit_price": unit_price,
            "tax": default_tax
        })

    if not rows:
        logger.warning(f"No valid rows parsed in {filename}")
        return

    result = {
        "delivery_date": delivery_date,
        "source_file": filename,
        "rows": rows
    }

    write_json_output(result, OUTPUT_DIR, filename, logger)
    print_last_log_lines(LOG_PATH, 10)

if __name__ == "__main__":
    with open("config/paths.json", encoding="utf-8") as f:
        path_config = json.load(f)["paths"]

    input_dir = path_config["input_cb"]
    for fname in os.listdir(input_dir):
        if fname.lower().endswith((".xlsx", ".xls")):
            parse_cb(os.path.join(input_dir, fname), path_config)
