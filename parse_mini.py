import re
import json
from pathlib import Path
import pdfplumber

# Load configs
with open("config/column_map_mini.json", "r", encoding="utf-8") as f:
    config = json.load(f)["mini_order"]

with open("config/paths.json", "r", encoding="utf-8") as f:
    path_cfg = json.load(f)["paths"]

input_folder = Path(path_cfg["input_mini"])
output_folder = Path(path_cfg["output_mini"])
output_folder.mkdir(parents=True, exist_ok=True)
log_path = output_folder / "parse_mini.log"

# Utility functions
def extract_date_from_text(text):
    match = re.search(r"(\d{1,2})-(\d{1,2})-(\d{4})", text)
    if match:
        day, month, year = match.groups()
        return f"{year}-{int(month):02d}-{int(day):02d}"
    return None

def extract_store_name(text):
    for line in text.splitlines():
        if "Store" in line:
            return line.replace("Store", "").strip()
    return "Unknown"

def to_int(s):
    try:
        return int(s.replace(",", ""))
    except:
        return 0

def to_float(s):
    try:
        return float(s.replace(",", ""))
    except:
        return 0.0

def safe_strip(val):
    return val.strip() if val else ""

# Parser using structured text
def parse_mini_text_pdf(pdf_path: Path) -> dict:
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text()

    delivery_date = extract_date_from_text(text)
    store = extract_store_name(text)
    rows = []

    for line in text.splitlines():
        match = re.match(r"^(\d{7})\s+(.+?)\s+EA\s+([\d,]+)\s+(\d+)\s+([\d,]+)", line)
        if match:
            sku, name, unit_price, qty, _ = match.groups()
            rows.append({
                "product_name": name,
                "qty": to_int(qty),
                "unit_price": to_float(unit_price),
                "tax": config["tax"]
            })

    return {
        "delivery_date": delivery_date,
        "store": store,
        "source_file": pdf_path.name,
        "rows": rows
    }

# Execution
logs = []

for pdf_file in input_folder.glob("*.pdf"):
    if "DONG XANH FOOD" not in pdf_file.name.upper():
        continue

    try:
        parsed = parse_mini_text_pdf(pdf_file)
        out_file = output_folder / pdf_file.with_suffix(".json").name
        with open(out_file, "w", encoding="utf-8") as f:
            json.dump(parsed, f, ensure_ascii=False, indent=2)
        logs.append(f"[OK] {pdf_file.name} → {len(parsed['rows'])} rows")
    except Exception as e:
        logs.append(f"[ERROR] {pdf_file.name}: {e}")

# Write log
with open(log_path, "w", encoding="utf-8") as f:
    for line in logs:
        f.write(line + "\n")

for line in logs[-10:]:
    print(line)
