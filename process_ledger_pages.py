import fitz
import os
import sys
import json
import time
import base64
import urllib.request
import urllib.error
from PIL import Image

# Configuration
PDF_PATH = "e:/user/OneDrive - Bangladesh Telecommunication Regulatory Commission/ABM/Sunnah/AL karim hisab/1. hisab alkarim 06-17-2026 23.19.pdf"
WORKSPACE_DIR = "e:/user/OneDrive - Bangladesh Telecommunication Regulatory Commission/ABM/Sunnah/AL karim hisab"
DATA_DIR = os.path.join(WORKSPACE_DIR, "data")
IMAGES_DIR = os.path.join(DATA_DIR, "images")
JSON_DIR = os.path.join(DATA_DIR, "json")

os.makedirs(IMAGES_DIR, exist_ok=True)
os.makedirs(JSON_DIR, exist_ok=True)

def load_api_key():
    env_path = "C:/Users/mannan/.env"
    if os.path.exists(env_path):
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                if line.strip().startswith("GEMINI_API_KEY="):
                    return line.strip().split("=", 1)[1].strip().strip('"').strip("'")
    return None

API_KEY = load_api_key()
if not API_KEY:
    print("Error: GEMINI_API_KEY not found")
    sys.exit(1)

def call_gemini(image_path, prompt, is_json=True, model="gemini-2.5-flash"):
    with open(image_path, "rb") as f:
        image_data = base64.b64encode(f.read()).decode("utf-8")
        
    payload = {
        "contents": [{
            "parts": [
                {"text": prompt},
                {
                    "inlineData": {
                        "mimeType": "image/png",
                        "data": image_data
                    }
                }
            ]
        }]
    }
    
    if is_json:
        payload["generationConfig"] = {"responseMimeType": "application/json"}
        
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={API_KEY}"
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"}
    )
    
    max_retries = 5
    backoff = 4
    for attempt in range(max_retries):
        try:
            with urllib.request.urlopen(req) as response:
                res = json.loads(response.read().decode("utf-8"))
                text = res["candidates"][0]["content"]["parts"][0]["text"]
                return text
        except urllib.error.HTTPError as e:
            if e.code == 429:
                print(f"Rate limit. Waiting {backoff}s (attempt {attempt+1}/{max_retries})...")
                time.sleep(backoff)
                backoff *= 2
            elif e.code >= 500:
                print(f"Server error {e.code}. Waiting {backoff}s...")
                time.sleep(backoff)
                backoff *= 2
            else:
                print(f"HTTP Error {e.code}: {e.read().decode('utf-8')}")
                raise e
        except Exception as e:
            print(f"Error: {e}. Retrying in {backoff}s...")
            time.sleep(backoff)
            backoff *= 2
            
    raise Exception("Max retries exceeded")

def detect_and_rotate_if_needed(pdf_page_num, page_width, page_height, img_path):
    # Check if page is portrait. If portrait, detect rotation.
    if page_width < page_height:
        print(f"Page {pdf_page_num} is portrait. Detecting rotation...")
        prompt = (
            "This is an image of a page from a handwritten Bengali ledger. It is currently oriented as portrait "
            "but the text itself might be rotated sideways. "
            "Determine the correct rotation (in degrees counter-clockwise: 0, 90, 180, 270) that we should apply to this image "
            "so the text becomes upright and readable horizontally from left to right. "
            "For example, if the text runs vertically from bottom to top (with the top of the page on the left), we need to rotate it 90 degrees clockwise, which is 270 degrees counter-clockwise. "
            "If the text runs vertically from top to bottom (with the top of the page on the right), we need to rotate it 90 degrees counter-clockwise. "
            "Return a JSON object with a single key 'rotation' containing the integer (0, 90, 180, or 270)."
        )
        try:
            res_text = call_gemini(img_path, prompt, is_json=True)
            data = json.loads(res_text)
            rotation = data.get("rotation", 0)
            if rotation in [90, 180, 270]:
                print(f"Rotating Page {pdf_page_num} by {rotation} degrees counter-clockwise...")
                img = Image.open(img_path)
                img = img.rotate(rotation, expand=True)
                img.save(img_path)
                print(f"Rotated and saved Page {pdf_page_num} image.")
        except Exception as e:
            print(f"Error detecting rotation for page {pdf_page_num}: {e}")
    else:
        print(f"Page {pdf_page_num} is landscape, assuming upright (no rotation).")

def process_ledger():
    doc = fitz.open(PDF_PATH)
    total_pages = len(doc)
    print(f"Total PDF pages: {total_pages}")
    
    # Process from page 3 to 110 (indices 2 to 109)
    for page_idx in range(2, total_pages):
        pdf_page_num = page_idx + 1
        json_path = os.path.join(JSON_DIR, f"page_{pdf_page_num}.json")
        img_path = os.path.join(IMAGES_DIR, f"page_{pdf_page_num}.png")
        
        if os.path.exists(json_path):
            print(f"Page {pdf_page_num} already processed. Skipping.")
            continue
            
        print(f"\n--- Processing Page {pdf_page_num}/{total_pages} ---")
        
        # Render image if not exists
        if not os.path.exists(img_path):
            page = doc[page_idx]
            pix = page.get_pixmap(dpi=150)
            pix.save(img_path)
            # Detect rotation and rotate
            detect_and_rotate_if_needed(pdf_page_num, page.rect.width, page.rect.height, img_path)
            
        prompt = (
            f"This is a photo of page {pdf_page_num} of a handwritten Bengali business ledger book. "
            "Please transcribe all the transaction records into structured JSON.\n\n"
            "Identify the following:\n"
            "1. The Client/Party name (নাম/পার্টি) and phone number, if any, written at the top.\n"
            "2. The Ledger Page Number(s) written at the top corners of the page.\n"
            "3. The debit (sales/bills) table (usually on the left side of a double-page spread, or the main table).\n"
            "   Columns: Date (তারিখ), Details (বিঃ কাঃ), Description (বিবরণ), Size (সাইজ), Model (মডেল), PD (পিঃ ডিঃ), Bill No (বিল), Qty (দর/পিস/ডঃ), Taka (টাকা), Total (মোট), Remarks (মন্তব্য).\n"
            "4. The credit (payments/cash received) table (usually on the right side of a double-page spread, or cash entries).\n"
            "   Columns: Date (তারিখ), Cash Amount (নগদ), Remarks (মন্তব্য).\n\n"
            "Return a JSON object with these exact keys:\n"
            "- 'party_name': 'string (transcribed from page, keep in Bengali)'\n"
            "- 'phone': 'string (transcribed from page, e.g. 017...)'\n"
            "- 'ledger_page_numbers': ['list of strings for page numbers at the top, e.g. [\"20\", \"21\"]']\n"
            "- 'debit_table': [\n"
            "    {\n"
            "      'no': 'string/number',\n"
            "      'date': 'string (convert Bengali digits to English digits, format DD.MM.YY)',\n"
            "      'bi_ka': 'string',\n"
            "      'description': 'string (keep in Bengali)',\n"
            "      'size': 'string (convert Bengali digits to English)',\n"
            "      'model': 'string (convert Bengali digits to English)',\n"
            "      'pd': 'string (convert Bengali digits to English)',\n"
            "      'bill': 'string (convert Bengali digits to English)',\n"
            "      'qty': 'string/number (convert Bengali digits to English)',\n"
            "      'taka': 'string/number (convert Bengali digits to English)',\n"
            "      'total': 'string/number (convert Bengali digits to English)',\n"
            "      'remarks': 'string (keep in Bengali)'\n"
            "    }\n"
            "  ],\n"
            "- 'credit_table': [\n"
            "    {\n"
            "      'no': 'string/number',\n"
            "      'date': 'string (convert Bengali digits to English digits, format DD.MM.YY)',\n"
            "      'amount': 'string/number (convert Bengali digits to English)',\n"
            "      'remarks': 'string (keep in Bengali)'\n"
            "    }\n"
            "  ]\n\n"
            "Notes:\n"
            "- Convert all numbers (like dates, quantities, rates, bills, totals, cash amounts) from Bengali digits to English digits (e.g. ১৭.৪.২৫ to 17.4.25, ৫২৬০ to 5260).\n"
            "- Do NOT translate any text or item descriptions, keep them in Bengali.\n"
            "- If a row or cell contains ditto marks or is blank but represents a continuation, you can inherit from the row above.\n"
            "- Ensure the JSON is valid and fits the schema exactly."
        )
        
        try:
            res_text = call_gemini(img_path, prompt, is_json=True)
            # Validate JSON
            data = json.loads(res_text)
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"Successfully saved Page {pdf_page_num} transcription.")
        except Exception as e:
            print(f"Failed to process Page {pdf_page_num}: {e}")
            
        # Rate limit safety delay
        time.sleep(3)

if __name__ == "__main__":
    process_ledger()
