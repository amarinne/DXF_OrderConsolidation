# parse_smile_cheers.py

import os
import json
import logging
import pandas as pd
from datetime import datetime

def setup_logger(log_path):
    os.makedirs(os.path.dirname(log_path), exist_ok=True)
    logger = logging.getLogger("smile_cheers_parser")
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

def parse_smile_cheers(file_path, path_config):
    filename = os.path.basename(file_path)

    with open("config/column_map_smile_cheers.json", encoding='utf-8') as f:
        config = json.load(f)

    LOG_PATH = os.path.join(path_config["output_smile_cheers"], "parse_smile_cheers.log")
    OUTPUT_DIR = path_config["output_smile_cheers"]
    logger = setup_logger(LOG_PATH)
    logger.info(f"Parsing file: {filename}")

    sheet_name = config["sheet_name"]
    header_row = config["header_row"] - 1
    col_map = config["columns"]
    tax = config["tax"]

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, engine="openpyxl")
        wb = pd.ExcelFile(file_path, engine="openpyxl")
        delivery_date_raw = wb.book[sheet_name][config["delivery_date_cell"]].value
        delivery_date = delivery_date_raw.strftime("%Y-%m-%d") if isinstance(delivery_date_raw, datetime) else None
    except Exception as e:
        logger.error(f"Failed to read file or delivery date: {e}")
        return

    rows = []
    for _, row in df.iterrows():
        product = str(row.get(col_map["product_name"])).strip()
        if not product or product.lower() in ["nan", "tổng cộng"]:
            continue

        try:
            qty = float(row.get(col_map["qty"]))
        except:
            continue

        try:
            unit_price = float(row.get(col_map["unit_price"]))
        except:
            unit_price = None

        rows.append({
            "product_name": product,
            "qty": qty,
            "unit_price": unit_price,
            "tax": tax
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

    input_dir = path_config["input_smile_cheers"]
    for fname in os.listdir(input_dir):
        if fname.lower().endswith((".xlsx", ".xls")):
            parse_smile_cheers(os.path.join(input_dir, fname), path_config)
