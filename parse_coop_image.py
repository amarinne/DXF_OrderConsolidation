import os
import re
import json
import logging
from datetime import datetime
import pytesseract
from pytesseract import Output
from PIL import Image
import cv2
import numpy as np

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

with open("config/paths.json", "r", encoding="utf-8") as f:
    PATHS = json.load(f)["paths"]

INPUT_DIR = PATHS["input_coop"]
OUTPUT_DIR = PATHS["output_coop"]

DATE_PATTERN = re.compile(r"(\d{1,2})[.](\d{1,2})")

def extract_date_from_filename(filename):
    match = DATE_PATTERN.search(filename)
    if not match:
        return None
    day, month = match.groups()
    try:
        date_obj = datetime.strptime(f"{day}.{month}.2025", "%d.%m.%Y")
        return date_obj.strftime("%Y-%m-%d")
    except ValueError:
        return None
def preprocess_image_for_ocr(image_path):
    image = cv2.imread(image_path)

    # Convert to HSV to mask red text
    hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
    lower_red1 = np.array([0, 50, 50])
    upper_red1 = np.array([10, 255, 255])
    lower_red2 = np.array([160, 50, 50])
    upper_red2 = np.array([180, 255, 255])
    red_mask = cv2.inRange(hsv, lower_red1, upper_red1) | cv2.inRange(hsv, lower_red2, upper_red2)

    # Force red text to black (OCR-readable)
    image[red_mask > 0] = [0, 0, 0]

    # Sharpen and threshold
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blur = cv2.GaussianBlur(gray, (3, 3), 0)
    sharp = cv2.addWeighted(gray, 1.5, blur, -0.5, 0)
    thresh = cv2.adaptiveThreshold(sharp, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY, 11, 2)
    return thresh

def parse_line_text(text):
    lines = [line.strip() for line in text.split("\n") if line.strip()]
    parsed_rows = []
    for line in lines:
        # Expecting format: name ... 5-digit price ... qty
        price_match = re.search(r"\b(\d{2,3}[.,]\d{3})\b", line)
        qty_match = re.search(r"\b(\d{1,3})\b(?=\D*$)", line)  # last number likely to be qty

        if price_match and qty_match:
            price_str = price_match.group(1).replace(",", "").replace(".", "")
            qty = float(qty_match.group(1))
            product_name = line[:price_match.start()].strip()
            parsed_rows.append({
                "product_name": product_name,
                "qty": qty,
                "unit_price": float(price_str),
                "tax": 0
            })
    return parsed_rows

def process_file(file_path):
    filename = os.path.basename(file_path)
    delivery_date = extract_date_from_filename(filename)

    image = preprocess_image_for_ocr(file_path)

    custom_oem_psm_config = r'--oem 3 --psm 6'
    text = pytesseract.image_to_string(image, lang='eng', config=custom_oem_psm_config)

    parsed_rows = parse_line_text(text)

    if not parsed_rows:
        logging.warning(f"No valid rows parsed from image: {filename}")
        return

    if filename.upper().startswith("DU KIEN"):
        label = "forecast"
    elif filename.upper().startswith("CHOT"):
        label = "confirmed"
    else:
        label = "unknown"

    output = {
        "delivery_date": delivery_date,
        "source_file": filename,
        "type": label,
        "rows": parsed_rows
    }

    output_file = os.path.join(OUTPUT_DIR, f"{os.path.splitext(filename)[0]}.json")
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    logging.info(f"Parsed and saved: {output_file}")

if __name__ == "__main__":
    for filename in os.listdir(INPUT_DIR):
        if filename.lower().endswith(('.jpg', '.jpeg', '.png')):
            process_file(os.path.join(INPUT_DIR, filename))
