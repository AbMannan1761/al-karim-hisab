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

# Load API key from C:/Users/mannan/.env
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
    print("Error: GEMINI_API_KEY not found in C:/Users/mannan/.env")
    sys.exit(1)

# Helper to call Gemini API
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
            # Handle rate limit (429) or server errors (5xx)
            if e.code == 429:
                print(f"Rate limit hit. Waiting {backoff} seconds (attempt {attempt+1}/{max_retries})...")
                time.sleep(backoff)
                backoff *= 2
            elif e.code >= 500:
                print(f"Server error {e.code}. Waiting {backoff} seconds (attempt {attempt+1}/{max_retries})...")
                time.sleep(backoff)
                backoff *= 2
            else:
                print(f"HTTP Error {e.code}: {e.read().decode('utf-8')}")
                raise e
        except Exception as e:
            print(f"Connection error: {e}. Retrying in {backoff} seconds...")
            time.sleep(backoff)
            backoff *= 2
            
    raise Exception("Max retries exceeded")

# Determine orientation of a portrait page
def get_portrait_rotation(image_path):
    prompt = (
        "This is an image of a page from a handwritten Bengali ledger book. The image might be rotated. "
        "Determine the correct rotation (in degrees counter-clockwise: 0, 90, 180, 270) to make the text upright "
        "and readable horizontally. Return a JSON object with a single key 'rotation' containing the integer value "
        "(0, 90, 180, or 270)."
    )
    try:
        res_text = call_gemini(image_path, prompt, is_json=True)
        data = json.loads(res_text)
        return data.get("rotation", 0)
    except Exception as e:
        print(f"Error detecting rotation for {image_path}: {e}")
        return 0

# Extract Index pages (PDF pages 1 and 2)
def process_index():
    index_file = os.path.join(JSON_DIR, "index_data.json")
    if os.path.exists(index_file):
        print("Index already processed. Loading from file...")
        with open(index_file, "r", encoding="utf-8") as f:
            return json.load(f)
            
    print("Processing Index pages...")
    doc = fitz.open(PDF_PATH)
    all_index_entries = []
    
    for page_idx in [0, 1]:
        pdf_page_num = page_idx + 1
        img_path = os.path.join(IMAGES_DIR, f"page_{pdf_page_num}.png")
        if not os.path.exists(img_path):
            page = doc[page_idx]
            pix = page.get_pixmap(dpi=150)
            pix.save(img_path)
            
        print(f"Transcribing Index Page {pdf_page_num}...")
        prompt = (
            "This is a photo of an index page (সূচীপত্র) from a handwritten Bengali ledger. "
            "Please extract all the rows from the table into a JSON array of objects.\n\n"
            "Each row should have these fields:\n"
            "- 'no': client number (নং) in English digits\n"
            "- 'party_name': name of the party (নাম-পার্টি)\n"
            "- 'page': ledger page number (পৃষ্ঠা) in English digits\n"
            "- 'address': address (ঠিকানা সমূহ)\n"
            "- 'notes': any other notes or marks written in the row\n\n"
            "Return a JSON array of these objects. Convert all numbers (like page numbers and client numbers) "
            "from Bengali digits to English digits (e.g. ০৬ to 6, ১২ to 12). Keep names and addresses in Bengali."
        )
        
        try:
            res_text = call_gemini(img_path, prompt, is_json=True)
            entries = json.loads(res_text)
            # Sometimes Gemini returns {"entries": [...]}, let's normalize
            if isinstance(entries, dict):
                for k in ["entries", "rows", "index"]:
                    if k in entries:
                        entries = entries[k]
                        break
            if isinstance(entries, list):
                all_index_entries.extend(entries)
                print(f"Extracted {len(entries)} entries from Index Page {pdf_page_num}.")
            else:
                print(f"Warning: Unexpected response format for Index Page {pdf_page_num}")
        except Exception as e:
            print(f"Failed to transcribe Index Page {pdf_page_num}: {e}")
            
    # Save index data
    with open(index_file, "w", encoding="utf-8") as f:
        json.dump(all_index_entries, f, ensure_ascii=False, indent=2)
    print(f"Saved total {len(all_index_entries)} index entries to {index_file}")
    return all_index_entries

if __name__ == "__main__":
    entries = process_index()
    print("Done testing index.")
