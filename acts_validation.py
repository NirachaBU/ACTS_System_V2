import os
import json
import requests
import pandas as pd
import time
from dotenv import load_dotenv
from google import genai
from pdf2image import convert_from_path

# --- 1. Setup ---
load_dotenv()
client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))
GEMINI_MODEL_ID = "models/gemini-2.0-flash" # แนะนำให้ใช้ v2.0-flash เพื่อความเสถียร

BASE_DIR = r"D:\ACTS_System_V2"
INPUT_PDF = os.path.join(BASE_DIR, "data", "raw_pdf", "test2_opal.pdf")
OUTPUT_PATH = os.path.join(BASE_DIR, "data", "output_excel", "ACTS_Master_Database_V2.xlsx")

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

def restructure_with_gemini(raw_text):
    print(f"[2/3] 🧠 Gemini Processing (Credit as Decimal 1.0)...")
    
    prompt = f"""
    สกัดข้อมูลจากทรานสคริปต์นี้เป็น JSON ภาษาไทย (ห้ามมีคำเกริ่น):
    {raw_text}

    โครงสร้าง JSON:
    {{
      "profile": {{
        "student_id": "รหัสนักศึกษา (string)",
        "name": "ชื่อ-นามสกุล (string)",
        "college": "ชื่อสถานศึกษา (string)",
        "graduated_level": "ระดับการศึกษาที่จบ เช่น ปวส. (string)",
        "faculty_type": "ประเภทวิชา (string)",
        "program_major": "สาขาวิชา (string)",
        "minor": "สาขางาน (string)",
        "entry_year": "ปีการศึกษาที่เข้าศึกษา (string)",
        "grad_year": "ปีการศึกษาที่สำเร็จการศึกษา (string)"
      }},
      "grades": [ 
        {{ 
          "code": "รหัสวิชา", 
          "subject": "ชื่อวิชาที่สะอาด", 
          "credit": 0.0, 
          "grade": "ผลการเรียน (เช่น 4.0 หรือ ผ.)" 
        }} 
      ],
      "on_paper_summary": {{ "total_credits": 0.0, "gpa": 0.0 }}
    }}

    กฎเหล็ก:
    1. **Credit Format**: หน่วยกิต (credit) ต้องเป็นตัวเลขทศนิยม 1 ตำแหน่งเสมอ (เช่น 3.0, 2.0, 1.0)
    2. **No Swapping**: ห้ามสลับช่องหน่วยกิตและผลการเรียน
    3. **Full Profile**: ดึงข้อมูลในส่วน profile มาให้ครบถ้วนทุกฟิลด์
    """
    
    try:
        response = client.models.generate_content(
            model=GEMINI_MODEL_ID,
            contents=prompt,
            config={'response_mime_type': 'application/json', 'temperature': 0.0}
        )
        return json.loads(response.text)
    except Exception as e:
        print(f" -> Gemini Error: {e}")
        return None

def save_to_master_excel(data, target_path):
    print(f"[3/3] 📊 Saving to Excel (Fixing Credit Data Type)...")
    try:
        os.makedirs(os.path.dirname(target_path), exist_ok=True)
        
        # 1. ชีท Students (ดึงค่ามาลงตาม DataType String)
        profile_df = pd.DataFrame([data['profile']])
        for col in profile_df.columns:
            profile_df[col] = profile_df[col].astype(str).replace('None', '')

        # 2. ชีท All_Grades
        grades_df = pd.DataFrame(data['grades'])
        grades_df = grades_df[grades_df['code'].notna() & (grades_df['code'] != "")]
        grades_df['student_id'] = data['profile']['student_id']
        
        # ปรับ Credit เป็นทศนิยม 1 ตำแหน่ง
        grades_df['credit'] = pd.to_numeric(grades_df['credit'], errors='coerce').fillna(0.0)
        grades_df['credit'] = grades_df['credit'].map('{:.1f}'.format) # บังคับ Format 1.0
        
        grades_df['grade'] = grades_df['grade'].astype(str).str.strip()

        # --- Validation Check ---
        scanned_credits = sum(pd.to_numeric(grades_df['credit']))
        paper_credits = float(data.get('on_paper_summary', {}).get('total_credits', 0.0))
        
        print("\n" + "="*45)
        print(f"👤 นักศึกษา: {data['profile']['name']}")
        print(f"📊 หน่วยกิตรวม (AI): {scanned_credits:.1f} | (บนกระดาษ): {paper_credits:.1f}")
        
        if abs(scanned_credits - paper_credits) < 0.1:
            print("✅ สถานะ: หน่วยกิตตรงกันเป๊ะ!")
        else:
            print(f"⚠️ คำเตือน: หน่วยกิตไม่ตรง! (ต่างกัน {abs(paper_credits - scanned_credits):.1f})")
        print("="*45 + "\n")

        with pd.ExcelWriter(target_path, engine='openpyxl') as writer:
            profile_df.to_excel(writer, sheet_name='Students', index=False)
            grades_df.to_excel(writer, sheet_name='All_Grades', index=False)
        
        return True
    except Exception as e:
        print(f" -> ❌ บันทึกพลาด: {e}")
        return False

if __name__ == "__main__":
    if os.path.exists(INPUT_PDF):
        raw_text = extract_text_with_typhoon(INPUT_PDF)
        if raw_text:
            final_json = restructure_with_gemini(raw_text)
            if final_json:
                save_to_master_excel(final_json, OUTPUT_PATH)
                print(f"✨ เสร็จเรียบร้อย! ข้อมูลมาครบ Credit เป็น 1.0 แล้ว")
                os.startfile(os.path.dirname(OUTPUT_PATH))
    else:
        print("❌ ไม่พบไฟล์ PDF")