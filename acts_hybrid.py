import os
import json
import re
import requests
import pandas as pd
import time
from dotenv import load_dotenv
from google import genai
from pdf2image import convert_from_path

# --- 1. Setup ---
load_dotenv()
client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))
GEMINI_MODEL_ID = "models/gemini-2.5-flash"

BASE_DIR = r"D:\ACTS_System_V2"
INPUT_PDF = os.path.join(BASE_DIR, "data", "raw_pdf", "test2_opal.pdf")
OUTPUT_PATH = os.path.join(BASE_DIR, "data", "output_excel", "ACTS_Master_Database_V2.xlsx")

# --- 2. OCR (คงเดิม) ---
def extract_text_with_typhoon(pdf_path):
    print(f"\n[1/3] 🌪️  Typhoon OCR Reading...")
    url = "https://api.opentyphoon.ai/v1/ocr"
    temp_image = "temp_processing.jpeg"
    try:
        pages = convert_from_path(pdf_path, dpi=300, first_page=1, last_page=1)
        if not pages: return None
        pages[0].save(temp_image, 'JPEG', quality=95)
        with open(temp_image, 'rb') as f:
            files = {'file': f}
            data = {'model': "typhoon-ocr", 'task_type': "default", 'max_tokens': 16384}
            headers = {'Authorization': f'Bearer {os.getenv("TYPHOON_API_KEY")}'}
            response = requests.post(url, files=files, data=data, headers=headers)
            if response.status_code == 200:
                content = response.json()['results'][0]['message']['choices'][0]['message']['content']
                try: return json.loads(content).get('natural_text', content)
                except: return content
            return None
    finally:
        if os.path.exists(temp_image): os.remove(temp_image)

# --- 3. ปรับ Prompt เพื่อดึงข้อมูลเพิ่มตามสั่ง ---
def restructure_with_gemini(raw_text):
    print(f"[2/3] 🧠 Gemini Processing (Full Student Profile Mode)...")
    
    prompt = f"""
    สกัดข้อมูลจากทรานสคริปต์นี้เป็น JSON ภาษาไทย:
    {raw_text}

    โครงสร้าง JSON:
    {{
      "profile": {{
        "student_id": "รหัสนักศึกษา (string)",
        "name": "ชื่อ-นามสกุล (string)",
        "college": "ชื่อสถานศึกษา (string)",
        "graduated_level": "ระดับการศึกษาที่จบ เช่น ปวส. (string)",
        "faculty_type": "ประเภทวิชา เช่น บริหารธุรกิจ (string ถ้ามี)",
        "program_major": "สาขาวิชา (string)",
        "minor": "สาขางาน (string)",
        "entry_year": "ปีการศึกษาที่เข้าศึกษา เช่น 2566 (string)",
        "grad_year": "ปีการศึกษาที่สำเร็จการศึกษา เช่น 2568 (string)"
      }},
      "grades": [ 
        {{ "code": "รหัสวิชา", "subject": "ชื่อวิชาที่สะอาด", "credit": 0, "grade": "string" }} 
      ]
    }}

    กฎการทำความสะอาด:
    - ชื่อวิชาต้องไม่มีคำว่า "ปีการศึกษา", "ภาคเรียน", หรือ "(*4)" ปนมา
    - สกัดเฉพาะบรรทัดที่มีรหัสวิชาเท่านั้น
    - เกรดตัวอักษร (ผ., ม.ส.) ให้คงไว้ตามจริง
    """
    
    try:
        response = client.models.generate_content(
            model=GEMINI_MODEL_ID,
            contents=prompt,
            config={'response_mime_type': 'application/json', 'temperature': 0.1}
        )
        return json.loads(response.text)
    except Exception as e:
        print(f" -> Gemini Error: {e}")
        return None

# --- 4. ปรับการเซฟและระบุ DataType ---
def save_to_master_excel(data, target_path):
    print(f"[3/3] 📊 Saving to Excel with Specific DataTypes...")
    try:
        os.makedirs(os.path.dirname(target_path), exist_ok=True)
        
        # 1. จัดการชีท Students และกำหนด DataType
        profile_df = pd.DataFrame([data['profile']])
        # บังคับ DataType เป็น String ทั้งหมดสำหรับ Profile เพื่อป้องกันข้อมูลเพี้ยน
        profile_cols = ['student_id', 'name', 'college', 'graduated_level', 
                        'faculty_type', 'program_major', 'minor', 'entry_year', 'grad_year']
        for col in profile_cols:
            if col in profile_df.columns:
                profile_df[col] = profile_df[col].astype(str).replace('None', '')

        # 2. จัดการชีท All_Grades
        grades_df = pd.DataFrame(data['grades'])
        grades_df = grades_df[grades_df['code'].notna() & (grades_df['code'] != "")]
        grades_df['student_id'] = data['profile']['student_id']
        
        # DataType สำหรับ Grades
        grades_df['code'] = grades_df['code'].astype(str)
        grades_df['subject'] = grades_df['subject'].astype(str)
        grades_df['credit'] = pd.to_numeric(grades_df['credit'], errors='coerce').fillna(0).astype(int)
        grades_df['grade'] = grades_df['grade'].astype(str).str.strip()

        with pd.ExcelWriter(target_path, engine='openpyxl') as writer:
            profile_df.to_excel(writer, sheet_name='Students', index=False)
            grades_df.to_excel(writer, sheet_name='All_Grades', index=False)
        
        print(f" -> ✅ บันทึกสำเร็จที่: {target_path}")
        return True
    except Exception as e:
        print(f" -> ❌ บันทึกพลาด: {e}")
        return False

if __name__ == "__main__":
    start_time = time.time()
    if os.path.exists(INPUT_PDF):
        raw_text = extract_text_with_typhoon(INPUT_PDF)
        if raw_text:
            final_json = restructure_with_gemini(raw_text)
            if final_json:
                save_to_master_excel(final_json, OUTPUT_PATH)
                print(f"\n✨ เสร็จสมบูรณ์ใน {round(time.time() - start_time, 2)} วินาที")
                os.startfile(os.path.dirname(OUTPUT_PATH))
    else:
        print("❌ ไม่พบไฟล์ PDF")