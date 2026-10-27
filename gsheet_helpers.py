# -*- coding: utf-8 -*-
"""
gsheet_helpers.py — เชื่อมต่อ Google Sheets/Drive ผ่าน Service Account
ใช้แทนสูตร IMPORTRANGE (ซึ่งทำงานได้เฉพาะใน Google Sheets) ด้วยการดึงค่าจริงมาคำนวณใน Python
และดึง STOCK.xlsx อัตโนมัติจาก Google Drive (พอร์ตจาก scripts/fetch_stock_drive.py)

ต้องตั้งค่า Service Account ก่อนใช้งาน (ดู README.md หัวข้อ "ตั้งค่า Google Service Account")
แล้วแชร์สิทธิ์ Viewer ของ Sheet/ไฟล์ที่เกี่ยวข้องให้กับอีเมลของ Service Account
"""
import io
import json

COST_SHEET_ID = "1whwXYiUyDC0uo6oaudr_QqzCV6UFOWd0mL4Y2El9JRY"
COST_SHEET_TAB = "CostORG2"
COST_RANGE = f"{COST_SHEET_TAB}!C3:N90000"   # col C = รหัสสินค้า, col N = ต้นทุน/ชิ้น (col ที่ 12 นับจาก C)

NOTE_SHEET_ID = "1p4-DX-Sq-Mvo_FNAkYtGMkdEBvPDCuPM_sZ_b-_Tm-Y"
NOTE_SHEET_TAB = "2026"
NOTE_RANGE = f"{NOTE_SHEET_TAB}!I2:Z29000"    # col I = เลขที่เอกสาร (key), col Z = หมายเหตุ (col ที่ 18 นับจาก I)

STOCK_DRIVE_FILE_ID = "1p4-DX-Sq-Mvo_FNAkYtGMkdEBvPDCuPM_sZ_b-_Tm-Y"
STOCK_SHEET_TAB = "2026"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive",  # ต้องเขียนไฟล์ได้ (ไม่ใช่แค่ readonly) เพื่ออัปโหลดผลลัพธ์ขึ้น Drive
]


def get_credentials(service_account_info: dict):
    """สร้าง google.oauth2 credentials จาก service account JSON (dict)"""
    from google.oauth2.service_account import Credentials
    return Credentials.from_service_account_info(service_account_info, scopes=SCOPES)


def get_user_drive_credentials(client_id: str, client_secret: str, refresh_token: str):
    """สร้าง credentials แบบ OAuth ผู้ใช้จริง (ไม่ใช่ Service Account) สำหรับอัปโหลดขึ้น Google Drive

    เหตุผลที่ต้องมีตัวนี้แยกจาก get_credentials(): Service Account ไม่มี storage quota เป็นของตัวเอง
    เลยสร้างไฟล์ใหม่ในโฟลเดอร์ Drive ส่วนตัว (บัญชี Gmail ธรรมดา ไม่ใช่ Shared Drive ของ Workspace) ไม่ได้
    -> จะเจอ Error 403 storageQuotaExceeded ทุกครั้ง
    เมื่อใช้ credentials ของผู้ใช้จริงแทน ไฟล์ที่อัปโหลดจะนับพื้นที่ของบัญชีผู้ใช้นั้น จึงอัปโหลดได้ปกติ

    วิธีได้ client_id / client_secret / refresh_token มา: ดู README.md หัวข้อ
    "ตั้งค่า OAuth สำหรับอัปโหลดขึ้น Google Drive อัตโนมัติ" (ใช้ OAuth Playground ครั้งเดียว)"""
    from google.oauth2.credentials import Credentials
    return Credentials(
        token=None,
        refresh_token=refresh_token,
        token_uri="https://oauth2.googleapis.com/token",
        client_id=client_id,
        client_secret=client_secret,
        scopes=["https://www.googleapis.com/auth/drive"],
    )


def _sheets_service(creds):
    from googleapiclient.discovery import build
    return build('sheets', 'v4', credentials=creds, cache_discovery=False)


def get_cost_map(creds, log=print) -> dict:
    """ดึงข้อมูลต้นทุนจาก Sheet CostORG2 -> {รหัสสินค้า: ต้นทุน/ชิ้น}
    แทนสูตร: VLOOKUP(code, IMPORTRANGE(COST_URL,"CostORG2!$C$3:$N$90000"), 12, 0)"""
    svc = _sheets_service(creds)
    result = svc.spreadsheets().values().get(
        spreadsheetId=COST_SHEET_ID, range=COST_RANGE
    ).execute()
    rows = result.get('values', [])
    cost_map = {}
    for row in rows:
        if not row or not row[0]:
            continue
        code = str(row[0]).strip()
        if len(row) >= 12 and row[11] not in (None, ''):
            try:
                cost_map[code] = float(str(row[11]).replace(',', ''))
            except ValueError:
                continue
    log(f"โหลดต้นทุนสินค้าจาก Google Sheets: {len(cost_map)} SKU")
    return cost_map


def get_note_map(creds, log=print) -> dict:
    """ดึงหมายเหตุจาก Sheet CSR 2026 -> {เลขที่เอกสาร: หมายเหตุ}
    แทนสูตร: INDEX(IMPORTRANGE(NOTE_URL,"2026!$I$2:$Z$29000"), MATCH(doc, ...col I...), 18)"""
    svc = _sheets_service(creds)
    result = svc.spreadsheets().values().get(
        spreadsheetId=NOTE_SHEET_ID, range=NOTE_RANGE
    ).execute()
    rows = result.get('values', [])
    note_map = {}
    for row in rows:
        if not row or not row[0]:
            continue
        key = str(row[0]).strip()
        note = row[17] if len(row) >= 18 else ''
        if key not in note_map:  # MATCH คืนค่าตัวแรกที่เจอ
            note_map[key] = note
    log(f"โหลดหมายเหตุจาก Google Sheets: {len(note_map)} เอกสาร")
    return note_map


def extract_folder_id(url_or_id: str) -> str:
    """รับได้ทั้ง Google Drive folder URL เต็มๆ หรือแค่ folder ID เปล่าๆ -> คืนแค่ folder ID
    เช่น https://drive.google.com/drive/folders/1AbC-XyZ?usp=sharing -> 1AbC-XyZ"""
    import re
    s = (url_or_id or "").strip()
    m = re.search(r"/folders/([a-zA-Z0-9_-]+)", s)
    if m:
        return m.group(1)
    m = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", s)
    if m:
        return m.group(1)
    return s  # สมมติว่า user วาง folder ID มาตรงๆ อยู่แล้ว


def upload_to_drive(creds, file_path: str, filename: str, folder_id: str, log=print,
                     convert_to_google_sheet: bool = True) -> dict:
    """อัปโหลดไฟล์ผลลัพธ์ (เช่น ค่าคอมLAZADA.xlsx) ขึ้น Google Drive folder ที่กำหนด
    ต้องแชร์สิทธิ์ Editor ของโฟลเดอร์นั้นให้กับอีเมลของ Service Account ก่อน

    convert_to_google_sheet=True (ค่าเริ่มต้น): สั่งให้ Google Drive แปลงไฟล์ .xlsx ที่อัปโหลด
    เป็น Google Sheets จริงทันที (ตั้ง mimeType ปลายทางเป็น 'application/vnd.google-apps.spreadsheet')
    -> สูตร IMPORTRANGE ในคอลัมน์ต้นทุน/กำไร/หมายเหตุ จะคำนวณค่าได้จริงทันทีที่เปิดไฟล์
    (ไฟล์ .xlsx ธรรมดาที่ดาวน์โหลดกลับเครื่องผ่านปุ่มดาวน์โหลด ยังคงเป็น Excel ปกติ ไม่แปลง)

    ⚠️ ครั้งแรกที่เปิดไฟล์ที่แปลงแล้ว Google Sheets อาจขึ้น popup "Allow access"
    ให้กดอนุญาตหนึ่งครั้ง (ผูกกับบัญชี Google ที่ใช้ตั้งค่า OAuth) หลังจากนั้นไฟล์ถัดๆ ไปที่อ้างอิง
    Sheet ต้นทางเดียวกันจะไม่ถามซ้ำอีก

    คืน dict {id, webViewLink} ของไฟล์ที่อัปโหลด"""
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload

    svc = build('drive', 'v3', credentials=creds, cache_discovery=False)
    folder_id = extract_folder_id(folder_id)
    if convert_to_google_sheet and filename.lower().endswith('.xlsx'):
        filename = filename[:-5]
    file_metadata = {'name': filename, 'parents': [folder_id]}
    if convert_to_google_sheet:
        file_metadata['mimeType'] = 'application/vnd.google-apps.spreadsheet'
    media = MediaFileUpload(
        file_path,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        resumable=True,
    )
    uploaded = svc.files().create(
        body=file_metadata, media_body=media, fields='id, webViewLink',
        supportsAllDrives=True,
    ).execute()
    kind = "Google Sheets (แปลงจาก .xlsx อัตโนมัติ)" if convert_to_google_sheet else "ไฟล์ .xlsx"
    log(f"อัปโหลดขึ้น Google Drive สำเร็จ: {filename} [{kind}] (file id: {uploaded.get('id')})")
    return uploaded


def fetch_stock_from_drive(creds, out_path: str, log=print) -> str:
    """ดึงชีท 2026 จาก Google Sheet STOCK มาเป็น STOCK.xlsx โดยตรง
    (เหมือนคัดลอกด้วยมือ Col A-U เท่านั้น ไม่ปรับโครงสร้าง — พอร์ตจาก fetch_stock_drive.py)"""
    from openpyxl import Workbook
    svc = _sheets_service(creds)
    result = svc.spreadsheets().values().get(
        spreadsheetId=STOCK_DRIVE_FILE_ID,
        range=f"{STOCK_SHEET_TAB}!A:U",
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()
    rows = result.get('values', [])

    wb = Workbook()
    ws = wb.active
    ws.title = 'STOCK'
    for row in rows:
        # pad ให้ครบ 21 คอลัมน์ (A-U) เหมือนต้นฉบับ
        padded = list(row) + [None] * (21 - len(row))
        ws.append(padded[:21])
    wb.save(out_path)
    log(f"STOCK saved: {out_path} ({max(len(rows)-1,0)} แถวข้อมูล + 1 แถว header, Col A-U)")
    return out_path
