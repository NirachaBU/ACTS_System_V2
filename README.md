# 🌪️ A.C.T.S. Hybrid System V2
ระบบสกัดข้อมูลทรานสคริปต์อาชีวะอัตโนมัติ (PDF to Excel) โดยใช้พลังของ **Typhoon OCR** และ **Gemini 2.5 Flash**

## ✨ Features
- **Dual Column Support**: รองรับทรานสคริปต์แบบ 2 คอลัมน์ (เทอม 1-2 ในหน้าเดียว)
- **Smart Grade Extraction**: เก็บเกรดได้ทั้งตัวเลข (4.0) และตัวอักษร (ผ., ม.ส., ม.ผ.)
- **Full Student Profile**: ดึงข้อมูลนักศึกษาครบถ้วน (รหัส, ชื่อ, สถาบัน, ปีที่จบ ฯลฯ)
- **Automated Excel**: จัดการข้อมูลแยกชีท Students และ Grades ให้อัตโนมัติ

## 🛠️ Installation
1. `pip install -r requirements.txt`
2. สร้างไฟล์ `.env` และใส่ API Keys
3. วางไฟล์ PDF ใน `data/raw_pdf/`
4. รัน `python acts_hybrid.py`
