import pdfplumber
import json
import re
from pathlib import Path

# Load column mapping
with open("config/column_map_genshai.json", "r", encoding="utf-8") as f:
    col_map = json.load(f)["genshai"]

# Load path config
with open("config/paths.json", "r", encoding="utf-8") as f:
    path_cfg = json.load(f)["paths"]

input_folder = Path(path_cfg["input_genshai"])
output_folder = Path(path_cfg["output_genshai"])
output_folder.mkdir(parents=True, exist_ok=True)
log_path = output_folder / "parse_genshai.log"

# Utilities
def extract_delivery_date(text):
    match = re.search(r"Ngày giao hàng:\s*(\d{1,2})/(\d{1,2})/(\d{4})", text)
    if match:
        day, month, year = match.groups()
        return f"{year}-{int(month):02d}-{int(day):02d}"
    return None

def to_int(val):
    try:
        return int(float(val.replace(",", "").strip()))
    except:
        return 0

def to_float(val):
    try:
        return float(val.replace(",", "").strip())
    except:
        return 0.0

def safe_strip(val):
    return val.strip() if val else ""

# Parser
def parse_genshai_pdf(pdf_path: Path) -> dict:
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text()
        delivery_date = extract_delivery_date(text)
        table = page.extract_table()
        rows = []

        if table:
            for row in table[1:]:
                if not row or len(row) <= max(col_map.values()):
                    continue
                if not safe_strip(row[col_map["product_name"]]):
                    continue
                rows.append({
                    "product_name": safe_strip(row[col_map["product_name"]]),
                    "qty": to_int(row[col_map["qty"]]),
                    "unit_price": to_float(row[col_map["unit_price"]]),
                    "tax": to_float(row[col_map["tax"]])
                })

    return {
        "delivery_date": delivery_date,
        "source_file": pdf_path.name,
        "rows": rows
    }

# Runner
logs = []

for pdf_file in input_folder.glob("*.pdf"):
    try:
        parsed_data = parse_genshai_pdf(pdf_file)
        output_path = output_folder / pdf_file.with_suffix(".json").name
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(parsed_data, f, ensure_ascii=False, indent=2)
        logs.append(f"[OK] {pdf_file.name} → {len(parsed_data['rows'])} rows")
    except Exception as e:
        logs.append(f"[ERROR] {pdf_file.name}: {e}")

# Log file output
with open(log_path, "w", encoding="utf-8") as f:
    for line in logs:
        f.write(line + "\n")

# Console preview
for line in logs[-10:]:
    print(line)
