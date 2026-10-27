# Lazada SKU Analysis — เว็บแอป

เว็บแอปที่พอร์ตมาจาก Claude Skill `lazada-sku-analysis` ให้ทีมงานอัปโหลดไฟล์ Lazada ดิบ
แล้วได้ไฟล์ `ค่าคอมLAZADA.xlsx` (3 ชีทรายงาน) กลับมาได้เอง ไม่ต้องเปิด Claude ทุกครั้ง

รองรับ 5 ร้าน: MAFIA, TRC, UTOPIA, FREEROAD, WORKFORCE

## โครงสร้างไฟล์

```
lazada-web-app/
├── streamlit_app.py       ← หน้าเว็บ (Streamlit)
├── lazada_core.py          ← Logic หลัก (พอร์ตจาก build_combined.py + build_analysis.py)
├── gsheet_helpers.py       ← เชื่อม Google Sheets/Drive (แทนสูตร IMPORTRANGE + ดึง STOCK อัตโนมัติ)
├── requirements.txt
├── .streamlit/secrets.toml.example
└── README.md
```

## 1) ทดสอบรันในเครื่องตัวเองก่อน (แนะนำ)

```bash
cd lazada-web-app
python -m venv venv && source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
streamlit run streamlit_app.py
```

เปิด `http://localhost:8501` แล้วลองอัปโหลดไฟล์ตัวอย่างดูก่อน ค่อย deploy ขึ้นคลาว

## 2) Deploy ขึ้นเว็บจริง (Streamlit Community Cloud — ฟรี ไม่ต้องใช้บัตรเครดิต)

1. สร้าง repo บน GitHub (private หรือ public ก็ได้) แล้ว push โฟลเดอร์ `lazada-web-app/` ทั้งหมดขึ้นไป
   ```bash
   git init
   git add .
   git commit -m "Lazada SKU Analysis web app"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<repo-name>.git
   git push -u origin main
   ```
2. ไปที่ **https://share.streamlit.io** แล้วล็อกอินด้วยบัญชี GitHub
3. กด **New app** → เลือก repo/branch → **Main file path** ใส่ `streamlit_app.py`
4. (ถ้าต้องการคำนวณต้นทุน/กำไร/หมายเหตุจริง หรือดึง STOCK อัตโนมัติ) เปิด **Advanced settings → Secrets**
   แล้ววางเนื้อหาจากขั้นตอนที่ 3 ด้านล่างลงไป
5. กด **Deploy** — รอ 1-2 นาที จะได้ลิงก์เว็บ เช่น `https://your-app.streamlit.app` ใช้งานได้จากที่ไหนก็ได้

> ต้องการให้ผมช่วยคลิกขั้นตอน GitHub/Streamlit Cloud ให้บนหน้าจอคุณเลยก็บอกได้ค่ะ (ผ่าน computer-use) แต่ขั้นตอนล็อกอินด้วยบัญชีจริงต้องให้คุณเป็นคนกดเอง

## 3) ตั้งค่า Google Service Account (ไม่บังคับ แต่แนะนำ)

ใช้สำหรับ 2 อย่าง:
- **คำนวณคอลัมน์ ต้นทุนสินค้า/กำไร/หมายเหตุ เป็นค่าจริง** แทนสูตร `IMPORTRANGE` (สูตรเดิมทำงานได้เฉพาะเปิดใน Google Sheets เท่านั้น เว็บนี้จะคำนวณให้เลยแทน)
- **ดึงไฟล์ STOCK.xlsx อัตโนมัติ** จาก Google Drive แทนการอัปโหลดเอง

ถ้าไม่ตั้งค่าก็ยังใช้งานได้ปกติ — เว็บจะเขียนสูตร IMPORTRANGE ไว้เหมือน Skill เดิม (ต้องอัปโหลดเข้า Google Sheets เองถึงจะเห็นค่า) และต้องอัปโหลด STOCK.xlsx เอง

### วิธีสร้าง Service Account

1. ไปที่ [Google Cloud Console](https://console.cloud.google.com/) → สร้างโปรเจกต์ใหม่ (หรือใช้ของเดิม)
2. เปิดใช้งาน API 2 ตัว: **Google Sheets API** และ **Google Drive API**
   (เมนู APIs & Services → Library → ค้นหาแล้วกด Enable)
3. ไปที่ **APIs & Services → Credentials → Create Credentials → Service Account**
   ตั้งชื่ออะไรก็ได้ เช่น `lazada-web-app`
4. เข้าไปที่ Service Account ที่สร้าง → แท็บ **Keys → Add Key → Create new key → JSON**
   จะได้ไฟล์ JSON ดาวน์โหลดมา (เก็บเป็นความลับ ห้าม commit ขึ้น GitHub)
5. เปิดไฟล์ JSON จะเห็นอีเมลแบบ `xxxx@xxxx.iam.gserviceaccount.com` — เอาอีเมลนี้ไป **แชร์สิทธิ์ Viewer**
   ให้กับ Google Sheets 2 ไฟล์นี้:
   - ต้นทุนสินค้า: `CostORG2` sheet (ลิงก์ที่ใช้ใน Skill เดิม)
   - หมายเหตุ + STOCK: `2026 CSRคุมสต๊อกOnline` (Google Drive file ID `1p4-DX-Sq-Mvo_FNAkYtGMkdEBvPDCuPM_sZ_b-_Tm-Y`)
6. เอาเนื้อหาในไฟล์ JSON มาใส่ในรูปแบบ `.streamlit/secrets.toml` (ดูตัวอย่างใน `secrets.toml.example`)
   - รันในเครื่อง: สร้างไฟล์ `.streamlit/secrets.toml` แล้ววางค่า
   - Deploy บนคลาว: วางค่าในหน้า **Advanced settings → Secrets** ตอนสร้าง app (หรือ App settings → Secrets ทีหลังได้)

## 4) บันทึกไฟล์ผลลัพธ์ขึ้น Google Drive อัตโนมัติ (ไม่บังคับ)

> ⚠️ **สำคัญ**: การแชร์สิทธิ์ Editor ให้ Service Account อย่างเดียว **ไม่พอ** ที่จะอัปโหลดไฟล์สำเร็จ เพราะ Service Account
> ไม่มี storage quota เป็นของตัวเอง (จะเจอ error `storageQuotaExceeded`) ต้องตั้งค่า OAuth เพิ่มตามข้อ 5) ด้วย
> ขั้นตอนด้านล่างนี้ (แชร์โฟลเดอร์ + folder ID) ยังต้องทำอยู่ แค่ "ผู้ได้รับสิทธิ์" เปลี่ยนจาก Service Account
> เป็นบัญชี Google ที่ใช้ authorize OAuth แทน

1. เตรียมโฟลเดอร์ปลายทางใน Google Drive ที่จะให้เว็บบันทึกไฟล์ลงไป
2. ถ้าบัญชี Google ที่จะใช้ authorize OAuth (ข้อ 5) เป็นเจ้าของโฟลเดอร์นี้อยู่แล้ว ข้ามขั้นตอนแชร์ได้เลย
   ไม่งั้นคลิกขวาที่โฟลเดอร์ → Share → ใส่อีเมลนั้น → ให้สิทธิ์ **Editor**
3. (ไม่บังคับ) ใส่ folder ID ไว้ล่วงหน้าใน secrets เพื่อไม่ต้องพิมพ์ทุกครั้ง — เพิ่มบรรทัดนี้ **ไว้นอก/เหนือ section อื่นๆ**
   ใน `.streamlit/secrets.toml` (สำคัญ: ต้องไม่อยู่ใต้ `[gcp_service_account]` หรือ section อื่น ไม่งั้นจะกลายเป็น key
   ที่ซ้อนอยู่ข้างในแทน อ่านค่าไม่ได้):
   ```toml
   drive_folder_id = "1AbC-XyZ..."   # เอาจากลิงก์โฟลเดอร์ .../folders/<ตรงนี้>

   [gcp_service_account]
   ...
   ```
4. ในหน้าเว็บ ติ๊ก "บันทึก ค่าคอมLAZADA.xlsx ขึ้น Google Drive อัตโนมัติ" แล้ววางลิงก์/Folder ID (หรือปล่อยว่างถ้าตั้งใน secrets แล้ว)
   — ช่องนี้จะขึ้นให้ติ๊กก็ต่อเมื่อตั้งค่า OAuth ตามข้อ 5) เสร็จแล้วเท่านั้น
5. ทุกครั้งที่ประมวลผลเสร็จ ระบบจะอัปโหลดไฟล์ชื่อ `ค่าคอมLAZADA_YYYYMMDD_HHMM.xlsx` ขึ้นโฟลเดอร์นั้นให้อัตโนมัติ พร้อมลิงก์เปิดไฟล์

## 5) ตั้งค่า OAuth สำหรับอัปโหลดขึ้น Google Drive อัตโนมัติ (จำเป็นถ้าอยากให้บันทึกขึ้น Drive จริงจัง)

**ทำไมต้องมีขั้นตอนนี้เพิ่ม:** Service Account (ข้อ 3) ไม่มี storage quota เป็นของตัวเอง ทำให้สร้างไฟล์ใหม่ในโฟลเดอร์ Drive
ส่วนตัว (บัญชี Gmail ธรรมดา ไม่ใช่ Shared Drive ของ Google Workspace) ไม่ได้ — จะเจอ error
`Service Accounts do not have storage quota` ทุกครั้ง ต้องใช้ credentials ของ "ผู้ใช้จริง" แทนตอนอัปโหลด
(ส่วนการอ่านข้อมูล Sheets/STOCK ในข้อ 3 ยังใช้ Service Account เหมือนเดิมได้ปกติ ไม่ต้องเปลี่ยน)

### ขั้นตอน (ทำครั้งเดียว)

1. **สร้าง OAuth Client ID** ใน [Google Cloud Console](https://console.cloud.google.com/apis/credentials)
   (โปรเจกต์เดียวกับที่สร้าง Service Account)
   - APIs & Services → Credentials → **Create Credentials → OAuth client ID**
   - Application type: **Web application**
   - Authorized redirect URIs: เพิ่ม `https://developers.google.com/oauthplayground`
   - กด Create → จะได้ **Client ID** และ **Client Secret** (เก็บไว้ ใช้ในขั้นตอนถัดไป)

2. **ตั้งค่า OAuth consent screen** (ถ้ายังไม่เคยตั้ง) — APIs & Services → OAuth consent screen
   - User Type: External
   - กรอกชื่อแอป/อีเมลติดต่อ (ใส่อะไรก็ได้ เพราะใช้ภายในทีมเอง)
   - Scopes: เพิ่ม `.../auth/drive`
   - Test users: เพิ่มอีเมล Google ที่จะใช้เป็นเจ้าของไฟล์บน Drive (เช่น `webmarketing01.trc@gmail.com`)
   - **สำคัญ**: กด **Publish App** ให้สถานะเป็น **In production** (ไม่ต้องส่งขอ verification ก็ได้ แค่กด publish)
     — ถ้าปล่อยไว้ที่ Testing, refresh token จะหมดอายุทุก 7 วัน ต้องขอใหม่เรื่อยๆ ถ้า publish แล้วจะใช้ได้ยาวๆ
     (ตอนเข้าเว็บ login ครั้งแรกจะเจอหน้าเตือน "Google hasn't verified this app" กดลิงก์ Advanced → ไปต่อได้ปกติ
     เพราะเป็นแอปที่เราสร้างเอง ไม่ใช่ของแปลกปลอม)

3. **ขอ Refresh Token ผ่าน OAuth Playground** — ไปที่ https://developers.google.com/oauthplayground
   - กดไอคอนเฟือง (⚙️) มุมขวาบน → ติ๊ก **Use your own OAuth credentials** → วาง Client ID และ Client Secret จากข้อ 1
   - Step 1: ค้นหา **Drive API v3** → เลือก scope `https://www.googleapis.com/auth/drive` → กด **Authorize APIs**
   - Login ด้วยบัญชี Google ที่จะเป็นเจ้าของไฟล์ (เช่น `webmarketing01.trc@gmail.com`) แล้วกด Allow
   - Step 2: กด **Exchange authorization code for tokens** → จะได้ **Refresh token** ก็อปเก็บไว้

4. **ใส่ค่าใน Secrets** (local: `.streamlit/secrets.toml`, cloud: Streamlit Cloud → App settings → Secrets) เพิ่มบล็อกนี้:
   ```toml
   [gdrive_oauth]
   client_id = "xxxxx.apps.googleusercontent.com"
   client_secret = "xxxxx"
   refresh_token = "1//xxxxx"
   ```

5. **แชร์สิทธิ์โฟลเดอร์ปลายทาง** (ข้อ 4) ให้บัญชีที่ใช้ authorize (ข้อ 3) เป็น Editor (ถ้าเป็นเจ้าของโฟลเดอร์อยู่แล้วข้ามได้)

6. รีบูตแอป (Streamlit Cloud → ⋮ → Reboot) แล้วติ๊ก "บันทึกขึ้น Google Drive" ในหน้าเว็บ — ตอนนี้จะอัปโหลดสำเร็จเพราะนับพื้นที่ของบัญชีผู้ใช้จริง ไม่ใช่ Service Account แล้ว

## หมายเหตุด้านความถูกต้อง (Parity กับ Skill เดิม)

- Logic การรวมไฟล์ + วิเคราะห์ทั้งหมดพอร์ตมาจาก `build_combined.py` และ `build_analysis.py` แบบตรงทุกจุด
  (ABB mapping, cross-month handling, กฎสีแดง col E/F, bundle logic, sort, color coding ฯลฯ)
- ส่วน "copy Data ERP พร้อม formatting" ในไฟล์ `build_combined.py` ต้นฉบับที่แถมมากับ Skill ถูกตัดขาดกลางคัน
  (ไฟล์จบกลางบรรทัด) ผมเติมโค้ดส่วนที่ขาดให้ครบตามที่ SKILL.md อธิบายไว้ (คัดลอก font/fill/border/alignment/number_format
  ทุก cell ด้วย openpyxl เพื่อรักษาไฮไลต์สีแดง) — แนะนำให้ทดสอบกับไฟล์จริงเทียบกับผลลัพธ์จาก Skill เดิมอีกครั้งก่อนใช้งานจริงจัง
- คอลัมน์ต้นทุน/กำไร/หมายเหตุ: ถ้าตั้งค่า Service Account จะได้ **ค่าตัวเลขจริง** (ดีกว่าสูตร IMPORTRANGE เพราะเปิดดูได้ทั้งใน
  Excel และ Google Sheets) ถ้าไม่ตั้งค่าจะได้สูตรเหมือนเดิมทุกประการ
- ยอดเงินเข้าบริษัทแต่ละร้านใน ชีท 1 ต้องตรงกับยอดรวมในชีท 3 เหมือน Skill เดิม — เว็บนี้ใช้สูตรคำนวณเดียวกันเป๊ะ ๆ
