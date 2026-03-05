import os
import json
import requests
import time
from dotenv import load_dotenv
from pdf2image import convert_from_path

# --- 1. ตั้งค่า API ---
load_dotenv()
TYPHOON_API_KEY = os.getenv("TYPHOON_API_KEY")

BASE_DIR = r"D:\ACTS_System_V2"
INPUT_PDF = os.path.join(BASE_DIR, "data", "raw_pdf", "test2_opal.pdf")
DEBUG_OUTPUT_TXT = os.path.join(BASE_DIR, "data", "typhoon_raw_output.txt")

def debug_typhoon_ocr(pdf_path):
    print("\n[DEBUG] 🌪️  กำลังส่งไฟล์ไปที่ Typhoon OCR...")
    url = "https://api.opentyphoon.ai/v1/ocr"
    temp_image = "temp_debug.jpeg"
    
    try:
        # แปลง PDF เป็นภาพ (เพิ่ม DPI เป็น 300 เพื่อความชัด จะได้อ่านมาครบๆ)
        pages = convert_from_path(pdf_path, dpi=300, first_page=1, last_page=1)
        if not pages:
            print("❌ แปลง PDF ไม่สำเร็จ")
            return
        pages[0].save(temp_image, 'JPEG')

        with open(temp_image, 'rb') as f:
            files = {'file': f}
            data = {
                'model': "typhoon-ocr", 
                'task_type': "default",
                'max_tokens': 16384  # ขอ Token เยอะๆ เพื่อให้ข้อความไม่โดนตัด
            }
            headers = {'Authorization': f'Bearer {TYPHOON_API_KEY}'}
            
            start_time = time.time()
            response = requests.post(url, files=files, data=data, headers=headers)
            duration = round(time.time() - start_time, 2)
            
            if response.status_code == 200:
                result = response.json()
                # ดึง Content ออกมา
                raw_content = result['results'][0]['message']['choices'][0]['message']['content']
                
                # พยายามแงะ natural_text (ถ้ามี)
                try:
                    parsed_json = json.loads(raw_content)
                    final_text = parsed_json.get('natural_text', raw_content)
                except:
                    final_text = raw_content

                print(f"✅ OCR สำเร็จ (ใช้เวลา {duration} วินาที)")
                print("\n" + "="*50)
                print("--- [ TYPHOON RAW OUTPUT START ] ---")
                print(final_text)
                print("--- [ TYPHOON RAW OUTPUT END ] ---")
                print("="*50)

                # บันทึกลงไฟล์ .txt เพื่อให้คุณไปนั่งไล่ดูได้ง่ายๆ
                with open(DEBUG_OUTPUT_TXT, "w", encoding="utf-8") as f:
                    f.write(final_text)
                
                print(f"\n📍 บันทึกข้อความดิบลงไฟล์แล้วที่: {DEBUG_OUTPUT_TXT}")
                os.startfile(DEBUG_OUTPUT_TXT) # สั่งเปิดไฟล์ txt ขึ้นมาเลย
                
            else:
                print(f"❌ Typhoon Error: {response.status_code} - {response.text}")

    finally:
        if os.path.exists(temp_image):
            os.remove(temp_image)

if __name__ == "__main__":
    if os.path.exists(INPUT_PDF):
        debug_typhoon_ocr(INPUT_PDF)
    else:
        print(f"❌ ไม่พบไฟล์ PDF ที่: {INPUT_PDF}")