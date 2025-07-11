# parse_satra.py

import os
import json
import logging
import openpyxl
import re

from datetime import datetime

def setup_logger(log_path):
    os.makedirs(os.path.dirname(log_path), exist_ok=True)
    logger = logging.getLogger("satra_parser")
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

def write_json_output(data, output_dir, base_name, suffix, logger):
    os.makedirs(output_dir, exist_ok=True)
    json_filename = f"{base_name}_{suffix}.json"
    output_path = os.path.join(output_dir, json_filename)
    try:
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        logger.info(f"Output written: {output_path}")
    except Exception as e:
        logger.error(f"Failed to write JSON output: {e}")

def parse_delivery_date(raw_value):
    if not raw_value:
        return None

    try:
        # Accept "dd.mm", "dd.mm.yy", "dd.mm.yyyy", "dd/mm", etc.
        parts = [int(p) for p in re.split(r"[.\-/]", str(raw_value).strip())]
        if len(parts) == 2:
            d, m = parts
            y = datetime.now().year
        elif len(parts) == 3:
            d, m, y = parts
            y = 2000 + y if y < 100 else y
        else:
            return None
        return datetime(y, m, d).strftime("%Y-%m-%d")
    except Exception:
        return None

def parse_satra(file_path, path_config):
    import re

    filename = os.path.basename(file_path)
    base_name = os.path.splitext(filename)[0]

    with open("config/column_map_satra.json", encoding='utf-8') as f:
        config = json.load(f)

    LOG_PATH = os.path.join(path_config["output_satra"], "parse_satra.log")
    OUTPUT_DIR = path_config["output_satra"]
    logger = setup_logger(LOG_PATH)
    logger.info(f"Parsing file: {filename}")

    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb.active
    except Exception as e:
        logger.error(f"Failed to open workbook: {e}")
        return

    delivery_date_raw = sheet[config["delivery_date_cell"]].value
    delivery_date = parse_delivery_date(delivery_date_raw)
    print(f"Raw delivery date value: {repr(delivery_date_raw)}")


    product_col = config["product_name_column"]
    start_row = config["product_name_header_row"] + 1
    tax = config["tax"]

    for warehouse, meta in config["warehouse_columns"].items():
        qty_col = meta["qty_col"]
        qty_start_row = meta["header_row"] + 1
        rows = []

        row_idx = max(start_row, qty_start_row)
        while True:
            product_cell = sheet[f"{product_col}{row_idx}"]
            qty_cell = sheet[f"{qty_col}{row_idx}"]

            product_name = str(product_cell.value).strip() if product_cell.value else ""
            if not product_name or product_name.lower() in ["tổng cộng", ""]:
                row_idx += 1
                continue

            try:
                qty = float(qty_cell.value)
            except (ValueError, TypeError):
                row_idx += 1
                continue

            rows.append({
                "product_name": product_name,
                "qty": qty,
                "unit_price": None,
                "tax": tax
            })
            row_idx += 1

            # Stop if all key cells in row are empty
            if all(sheet[f"{c}{row_idx}"].value in [None, "", " "] for c in [product_col, qty_col]):
                break

        if rows:
            result = {
                "delivery_date": delivery_date,
                "source_file": filename,
                "store": warehouse,
                "rows": rows
            }
            write_json_output(result, OUTPUT_DIR, base_name, warehouse.lower(), logger)

    print_last_log_lines(LOG_PATH, 10)

if __name__ == "__main__":
    with open("config/paths.json", encoding="utf-8") as f:
        path_config = json.load(f)["paths"]

    input_dir = path_config["input_satra"]
    for fname in os.listdir(input_dir):
        if fname.lower().endswith((".xlsx", ".xls")):
            parse_satra(os.path.join(input_dir, fname), path_config)
