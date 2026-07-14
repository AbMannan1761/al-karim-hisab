import fitz
import os
import sys
import io
import json
import time
import base64
import urllib.request
import urllib.error
from PIL import Image

try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass


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

# Load index mapping
index_file = os.path.join(JSON_DIR, "index_data.json")
index_map = {}
if os.path.exists(index_file):
    with open(index_file, "r", encoding="utf-8") as f:
        idx_entries = json.load(f)
        for entry in idx_entries:
            try:
                p_num = int(entry.get("page", 0))
                if p_num:
                    index_map[p_num] = {
                        "name": entry.get("party_name", ""),
                        "address": entry.get("address", ""),
                        "notes": entry.get("notes", "")
                    }
            except Exception:
                pass

# Fallback models list to maximize daily quota utilization (each model has 20 RPD on free tier)
MODELS = [
    "gemini-3.5-flash",
    "gemini-3.1-flash-lite",
    "gemini-3.1-flash-lite-preview",
    "gemini-flash-lite-latest",
    "gemini-3-flash-preview",
    "gemini-2.5-flash-lite",
    "gemini-2.5-flash",
    "gemini-flash-latest"
]
current_model_idx = 0

def call_gemini(image_path, prompt, is_json=True):
    global current_model_idx
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
        
    max_retries = 5
    backoff = 4
    
    while current_model_idx < len(MODELS):
        model = MODELS[current_model_idx]
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={API_KEY}"
        req = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"}
        )
        
        attempt = 0
        model_retry = True
        while model_retry and attempt < max_retries:
            try:
                with urllib.request.urlopen(req, timeout=60) as response:
                    res = json.loads(response.read().decode("utf-8"))
                    text = res["candidates"][0]["content"]["parts"][0]["text"]
                    return text
            except urllib.error.HTTPError as e:
                err_body = ""
                try:
                    err_body = e.read().decode("utf-8")
                except Exception:
                    pass
                
                if e.code == 429:
                    if "quota" in err_body.lower() and ("limit: 20" in err_body.lower() or "limit: 0" in err_body.lower() or "exceeded your current quota" in err_body.lower() or "free_tier_requests" in err_body.lower()):
                        print(f"Daily quota exceeded for {model}. Switching to next model...")
                        current_model_idx += 1
                        model_retry = False
                    else:
                        print(f"Rate limit (RPM/TPM) for {model}. Waiting {backoff}s (attempt {attempt+1}/{max_retries})...")
                        time.sleep(backoff)
                        backoff *= 2
                        attempt += 1
                elif e.code == 503:
                    print(f"Server error 503 for {model}. Waiting {backoff}s (attempt {attempt+1}/{max_retries})...")
                    time.sleep(backoff)
                    backoff *= 2
                    attempt += 1
                elif e.code >= 500:
                    print(f"Server error {e.code} for {model}. Waiting {backoff}s (attempt {attempt+1}/{max_retries})...")
                    time.sleep(backoff)
                    backoff *= 2
                    attempt += 1
                else:
                    print(f"HTTP Error {e.code} for {model}: {err_body}")
                    # Switch model for other non-transient errors
                    current_model_idx += 1
                    model_retry = False
            except Exception as e:
                print(f"Error for {model}: {e}. Retrying in {backoff}s...")
                time.sleep(backoff)
                backoff *= 2
                attempt += 1
                
        if attempt >= max_retries:
            print(f"Max retries exceeded for {model}. Switching to next model...")
            current_model_idx += 1
            
    raise Exception("All models exhausted or failed.")

def detect_and_rotate_if_portrait(pdf_page_num, page_width, page_height, img_path):
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
                print(f"Rotating Page {pdf_page_num} by {rotation} degrees counter-clockwise to make it landscape...")
                img = Image.open(img_path)
                img = img.rotate(rotation, expand=True)
                img.save(img_path)
                print(f"Rotated and saved Page {pdf_page_num} image.")
        except Exception as e:
            print(f"Error detecting rotation for page {pdf_page_num}: {e}")
    else:
        print(f"Page {pdf_page_num} is landscape, assuming upright.")

def process_single_half(img_path, pdf_page_num, side, expected_client, expected_page):
    json_path = os.path.join(JSON_DIR, f"page_{pdf_page_num}_{side}.json")
    if os.path.exists(json_path):
        print(f"  {side.capitalize()} half already processed.")
        return
        
    print(f"  Processing {side.capitalize()} half...")
    
    hint_text = ""
    if expected_client:
        hint_text = (
            f"According to the ledger index, the client/party on ledger page {expected_page} is expected to be: '{expected_client}'. "
            "Please check the top of the page and transcribe the Client/Party name and phone number. "
            "Use this expected client name as a strong hint, but read the handwritten characters on the page carefully to confirm."
        )
        
    if side == "left":
        prompt = (
            f"This is a photo of the left page of a handwritten Bengali ledger sheet. This page represents the Debit (Sales/Bills) entries. "
            f"Ledger Page Number should be around {expected_page}. {hint_text}\n\n"
            "Please transcribe all the sales/debit transaction rows from the table into a JSON object.\n\n"
            "The columns on the page are:\n"
            "No (নং), Date (তারিখ), Details (বিঃ কাঃ), Description (বিবরণ), Size (সাইজ), Model (মডেল), Qty (পিস), Bill No (বিল), Rate (দর), Amount (টাকা), Total (মোট), Remarks (মন্তব্য).\n\n"
            "Return a JSON object with these exact keys:\n"
            "- 'party_name': 'string (transcribed name in Bengali)'\n"
            "- 'phone': 'string (e.g., 017...)'\n"
            "- 'ledger_page_number': 'string (ledger page number, e.g., 8)'\n"
            "- 'debit_table': [\n"
            "    {\n"
            "      'no': 'string/number',\n"
            "      'date': 'string (convert Bengali digits to English digits, format DD.MM.YY)',\n"
            "      'bi_ka': 'string',\n"
            "      'description': 'string (keep in Bengali)',\n"
            "      'size': 'string (convert Bengali digits to English)',\n"
            "      'model': 'string (convert Bengali digits to English)',\n"
            "      'pd': 'string (transcribe values from the \"টাকা\" column here, convert Bengali digits to English)',\n"
            "      'bill': 'string (transcribe values from the \"বিল\" column here, convert Bengali digits to English)',\n"
            "      'qty': 'string/number (transcribe values from the \"পিস\" column here, convert Bengali digits to English)',\n"
            "      'taka': 'string/number (transcribe values from the \"দর\" column here, convert Bengali digits to English)',\n"
            "      'total': 'string/number (transcribe values from the \"মোট\" column here, convert Bengali digits to English)',\n"
            "      'remarks': 'string (keep in Bengali)'\n"
            "    }\n"
            "  ]\n\n"
            "Notes:\n"
            "- Convert all numbers (like dates, quantities, rates, bills, totals) from Bengali digits to English digits (e.g. ১৭.৪.২৫ to 17.4.25, ৫২৬০ to 5260).\n"
            "- Do NOT translate any text or item descriptions, keep them in Bengali.\n"
            "- If a row or cell contains ditto marks (\") or is blank but represents a continuation, you can inherit from the row above.\n"
            "- Ensure the JSON is valid."
        )
    else:
        prompt = (
            f"This is a photo of the right page of a handwritten Bengali ledger sheet. This page represents the Credit (Payments/Cash Received) entries. "
            f"Ledger Page Number should be around {expected_page}. {hint_text}\n\n"
            "Please transcribe all the credit/payment transaction rows from the table into a JSON object.\n\n"
            "The columns are:\n"
            "No (নং), Date (তারিখ), Cash Amount (নগদ), Remarks (মন্তব্য).\n\n"
            "Return a JSON object with these exact keys:\n"
            "- 'party_name': 'string (transcribed name in Bengali)'\n"
            "- 'phone': 'string (e.g., 017...)'\n"
            "- 'ledger_page_number': 'string (ledger page number, e.g., 9)'\n"
            "- 'credit_table': [\n"
            "    {\n"
            "      'no': 'string/number',\n"
            "      'date': 'string (convert Bengali digits to English digits, format DD.MM.YY)',\n"
            "      'amount': 'string/number (convert Bengali digits to English)',\n"
            "      'remarks': 'string (keep in Bengali)'\n"
            "    }\n"
            "  ]\n\n"
            "Notes:\n"
            "- Convert all numbers (like dates, cash amounts) from Bengali digits to English digits (e.g. ১৭.৪.২৫ to 17.4.25, ৫০০০ to 5000).\n"
            "- Do NOT translate any text, keep Bengali text in Bengali.\n"
            "- Ensure the JSON is valid."
        )
        
    for parse_attempt in range(3):
        try:
            res_text = call_gemini(img_path, prompt, is_json=True)
            data = json.loads(res_text)
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"    Saved {side} transcription.")
            break
        except json.JSONDecodeError as je:
            print(f"    JSON decode error on attempt {parse_attempt+1}: {je}. Retrying JSON generation...")
            time.sleep(2)
        except Exception as e:
            print(f"    Failed to process {side} half of page {pdf_page_num}: {e}")
            break
        
    # Small delay between API calls
    time.sleep(3)

def process_ledger():
    doc = fitz.open(PDF_PATH)
    total_pages = len(doc)
    print(f"Total PDF pages: {total_pages}")
    
    # Process page 3 to 110
    for page_idx in range(2, total_pages):
        pdf_page_num = page_idx + 1
        
        # Define files
        img_path = os.path.join(IMAGES_DIR, f"page_{pdf_page_num}.png")
        left_img_path = os.path.join(IMAGES_DIR, f"page_{pdf_page_num}_left.png")
        right_img_path = os.path.join(IMAGES_DIR, f"page_{pdf_page_num}_right.png")
        
        left_json = os.path.join(JSON_DIR, f"page_{pdf_page_num}_left.json")
        right_json = os.path.join(JSON_DIR, f"page_{pdf_page_num}_right.json")
        
        # Check if already processed both halves
        if os.path.exists(left_json) and os.path.exists(right_json):
            print(f"Page {pdf_page_num} already processed. Skipping.")
            continue
            
        print(f"\n--- Processing Page {pdf_page_num}/{total_pages} ---")
        
        # Render page
        if not os.path.exists(img_path):
            page = doc[page_idx]
            pix = page.get_pixmap(dpi=150)
            pix.save(img_path)
            # Check rotation
            detect_and_rotate_if_portrait(pdf_page_num, page.rect.width, page.rect.height, img_path)
            
        # Ensure split images are saved
        if not os.path.exists(left_img_path) or not os.path.exists(right_img_path):
            try:
                img = Image.open(img_path)
            except Exception as e:
                print(f"  Failed to open image {img_path}: {e}. Re-rendering page...")
                if os.path.exists(img_path):
                    try:
                        os.remove(img_path)
                    except Exception:
                        pass
                page = doc[page_idx]
                pix = page.get_pixmap(dpi=150)
                pix.save(img_path)
                detect_and_rotate_if_portrait(pdf_page_num, page.rect.width, page.rect.height, img_path)
                img = Image.open(img_path)
                
            w, h = img.size
            
            # Split based on binder ratio (Debit page is wide, Credit page is narrow)
            split_x = int(w * 0.77)
            left_img = img.crop((0, 0, split_x, h))
            left_img.save(left_img_path)
            
            right_img = img.crop((split_x, 0, w, h))
            right_img.save(right_img_path)
            print(f"  Split page {pdf_page_num} into left and right images at 77% ratio.")
            
        # Expected page numbers: Left is 2 * pdf_page_num, Right is 2 * pdf_page_num + 1
        left_ledger_page = 2 * pdf_page_num
        right_ledger_page = 2 * pdf_page_num + 1
        
        # Find expected client name
        # We look up in index_map using left_ledger_page
        expected_client_info = index_map.get(left_ledger_page)
        expected_client_name = expected_client_info["name"] if expected_client_info else None
        
        # If left page info not found, check right page (just in case)
        if not expected_client_name:
            expected_client_info = index_map.get(right_ledger_page)
            expected_client_name = expected_client_info["name"] if expected_client_info else None
            
        print(f"  Expected client: {expected_client_name} (Ledger pages {left_ledger_page}-{right_ledger_page})")
        
        # Process left half (Debit)
        process_single_half(left_img_path, pdf_page_num, "left", expected_client_name, left_ledger_page)
        
        # Process right half (Credit)
        process_single_half(right_img_path, pdf_page_num, "right", expected_client_name, right_ledger_page)

if __name__ == "__main__":
    process_ledger()
